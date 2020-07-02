VERSION 5.00
Begin VB.Form GDQuickButtonfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Button Options"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   2295
   Icon            =   "GDQuickButtonfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   2295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save option changes"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Frame frmQuickOptions 
      Caption         =   "Check items for pasting"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.CheckBox chkSampleSources 
         Caption         =   "Sample Sources"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         ToolTipText     =   "Click to paste whether Outcropping/Well (also which type)"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkDepths 
         Caption         =   "Depths"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         ToolTipText     =   "Check to paste depths"
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox chkGroundLevels 
         Caption         =   "Ground Levels"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Check to Paste Ground Levels"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkGeologicAges 
         Caption         =   "Geologic Ages"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "Check to paste geologic ages"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox chkFormations 
         Caption         =   "Formations"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Check to paste formations"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkFossils 
         Caption         =   "Fossil Types"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         ToolTipText     =   "Check to paste fossil types"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox chkCoordinates 
         Caption         =   "Coordinates"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Check to paste coordinates"
         Top             =   1280
         Width           =   1215
      End
      Begin VB.CheckBox chkNames 
         Caption         =   "Names"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Check to paste names"
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "GDQuickButtonfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()

With GDQuickButtonfrm

   If .chkNames.value = vbChecked Then
      bNames = True
   Else
      bNames = False
      End If
      
   If .chkGroundLevels.value = vbChecked Then
      bGroundLevels = True
   Else
      bGroundLevels = False
      End If
      
   If .chkDepths.value = vbChecked Then
      bDepths = True
   Else
      bDepths = False
      End If
      
   If .chkFossils.value = vbChecked Then
      bFossilTypes = True
   Else
      bFossilTypes = False
      End If
      
   If .chkCoordinates.value = vbChecked Then
      bCoordinates = True
   Else
      bCoordinates = False
      End If
      
   If .chkGeologicAges.value = vbChecked Then
      bGeologicAges = True
   Else
      bGeologicAges = False
      End If
      
   If .chkSampleSources.value = vbChecked Then
      bSampleSources = True
   Else
      bSampleSources = False
      End If
      
   If .chkFormations.value = vbChecked Then
      bFormations = True
   Else
      bFormations = False
      End If
      
   optionfil% = FreeFile
   Open direct$ & "\quickbuttonoptions.txt" For Output As #optionfil%
   
   Write #optionfil%, "This file is used by the MapDigitizer program, don't erase it! "
   Write #optionfil%, "Scanned database quickbutton template, revised: " & Date

   Write #optionfil%, bNames
   Write #optionfil%, bGroundLevels
   Write #optionfil%, bDepths
   Write #optionfil%, bFossilTypes
   Write #optionfil%, bCoordinates
   Write #optionfil%, bGeologicAges
   Write #optionfil%, bSampleSources
   Write #optionfil%, bFormations
   Close #optionfil%
   
   Unload Me
   
End With
End Sub

Private Sub Form_Load()

   On Error GoTo Form_Load_Error

   QuickVis = True

   If Dir(direct$ & "\quickbuttonoptions.txt") <> sEmpty Then
      
      optionfil% = FreeFile
      Open direct$ & "\quickbuttonoptions.txt" For Input As #optionfil%
      
      Line Input #optionfil%, doclin$
      Line Input #optionfil%, doclin$
      
      Input #optionfil%, bNames
      Input #optionfil%, bGroundLevels
      Input #optionfil%, bDepths
      Input #optionfil%, bFossilTypes
      Input #optionfil%, bCoordinates
      Input #optionfil%, bGeologicAges
      Input #optionfil%, bSampleSources
      Input #optionfil%, bFormations
      Close #optionfil%
      
      End If
      
   If bNames Then
      chkNames.value = vbChecked
   Else
      chkNames.value = vbUnchecked
      End If
      
   If bGroundLevels Then
      chkGroundLevels.value = vbChecked
   Else
      chkGroundLevels.value = vbUnchecked
      End If
      
   If bDepths Then
      chkDepths.value = vbChecked
   Else
      chkDepths.value = vbUnchecked
      End If
      
   If bFossilTypes Then
      chkFossils.value = vbChecked
   Else
      chkDepths.value = vbUnchecked
      End If
      
   If bCoordinates Then
      chkCoordinates.value = vbChecked
   Else
      chkCoordinates.value = vbUnchecked
      End If
      
   If bGeologicAges Then
      chkGeologicAges.value = vbChecked
   Else
      chkGeologicAges.value = vbUnchecked
      End If
      
   If bSampleSources Then
      chkSampleSources.value = vbChecked
   Else
      chkSampleSources.value = vbUnchecked
      End If
      
   If bFormations Then
      chkFormations.value = vbChecked
   Else
      chkFormations.value = vunbChecked
      End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

   Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
   QuickVis = False
   Unload Me
   Set GDQuickButtonfrm = Nothing
End Sub
