VERSION 5.00
Begin VB.Form Caldirectories 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cal Program"
   ClientHeight    =   8490
   ClientLeft      =   3570
   ClientTop       =   1920
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cal Progarm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   4560
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Kovitz list"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3200
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   40
      ToolTipText     =   "Use Kovitz sorted list"
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame frmTableType 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   430
      Left            =   360
      TabIndex        =   36
      Top             =   2620
      Width           =   3855
      Begin VB.OptionButton optBoth 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Both"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   160
         Width           =   975
      End
      Begin VB.OptionButton optSunset 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sunset"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   38
         Top             =   160
         Width           =   1095
      End
      Begin VB.OptionButton optSunrise 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sunrise"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   37
         Top             =   160
         Width           =   1095
      End
   End
   Begin VB.TextBox txtCivil 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   29
      ToolTipText     =   "enter directories suffix, e.g. ""2012-2015"""
      Top             =   1830
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkOnePage 
      BackColor       =   &H00FFFFC0&
      Caption         =   "one page only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   34
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox chkPdfPrinter 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pdf printer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1900
      TabIndex        =   33
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame frmRounding 
      BackColor       =   &H00FFFFC0&
      Height          =   500
      Left            =   360
      TabIndex        =   30
      Top             =   2160
      Width           =   3855
      Begin VB.ComboBox cmbRounding 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2000
         TabIndex        =   32
         Top             =   150
         Width           =   1575
      End
      Begin VB.CheckBox chkAutoRounding 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Set rounding"
         Enabled         =   0   'False
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   200
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkListAuto 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Save as html list"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   480
      TabIndex        =   28
      Top             =   1360
      Width           =   2415
   End
   Begin VB.Frame frmYearType 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   360
      TabIndex        =   25
      Top             =   1620
      Width           =   975
      Begin VB.OptionButton optEnglish 
         BackColor       =   &H00FFFFC0&
         Caption         =   "civil"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   120
         TabIndex        =   27
         Top             =   330
         Width           =   745
      End
      Begin VB.OptionButton optHebrew 
         BackColor       =   &H00FFFFC0&
         Caption         =   "hebrew"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   130
         Value           =   -1  'True
         Width           =   745
      End
   End
   Begin VB.Frame frmCalType 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   3120
      TabIndex        =   21
      ToolTipText     =   "Which type of tables to calculate"
      Top             =   800
      Width           =   1095
      Begin VB.CheckBox chkObs 
         Enabled         =   0   'False
         Height          =   240
         Left            =   720
         TabIndex        =   35
         ToolTipText     =   "Use added cushion option"
         Top             =   130
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton optAst 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ast"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   550
         Width           =   855
      End
      Begin VB.OptionButton optMishor 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mishor"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   330
         Width           =   855
      End
      Begin VB.OptionButton optVis 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Vis"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   130
         Width           =   855
      End
   End
   Begin VB.CheckBox chkHtmlAuto 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Save as  html"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   1150
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox chkPrintAuto 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&default Printer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   950
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Visual sunrise times using the GTOPO10"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Check for near mtns and display in different. color"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   360
      TabIndex        =   17
      ToolTipText     =   "Displays dubious times in a different color"
      Top             =   600
      Width           =   3795
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Use sorted city list citysort.txt"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   360
      TabIndex        =   16
      ToolTipText     =   "Prints messages if detects alot of near obstructions"
      Top             =   360
      Value           =   -1  'True
      Width           =   3855
   End
   Begin VB.CheckBox Astroncheck 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Calculate Sunrise/Sunset tables for user inputed coordinates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   15
      ToolTipText     =   "User input of coordinates"
      Top             =   2640
      Visible         =   0   'False
      Width           =   3960
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   1560
      TabIndex        =   13
      Text            =   "1"
      Top             =   3140
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Paginate/Table of Contents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2760
      TabIndex        =   12
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton AutoCancelbut 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Picture         =   "Cal Progarm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   110
      Width           =   1215
   End
   Begin VB.CommandButton Runbutton 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Begin         Run "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Picture         =   "Cal Progarm.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1770
      Width           =   855
   End
   Begin VB.CommandButton Autobut 
      BackColor       =   &H00FFFF80&
      Caption         =   "Enable automatic operation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   110
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   4095
   End
   Begin VB.CommandButton ExitButton 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Aharoni"
         Size            =   20.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      Picture         =   "Cal Progarm.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   4095
   End
   Begin VB.CommandButton OKbutton 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Aharoni"
         Size            =   26.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Picture         =   "Cal Progarm.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "c:\cities\jerusalem_other_neighbohoods"
      Top             =   7080
      Width           =   4095
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808000&
      Height          =   705
      Left            =   300
      Top             =   915
      Width           =   2715
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808000&
      Height          =   495
      Left            =   300
      Top             =   360
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      Height          =   330
      Left            =   240
      Top             =   1920
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Beginning Page Num. of Tables  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      Height          =   3495
      Left            =   240
      Top             =   75
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1360
      TabIndex        =   9
      Top             =   1830
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Chosen Directory (city name)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pick Directory and Enter OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   4095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Caldirectories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Autobut_Click()
   Label3.Enabled = True
   Combo1.Enabled = True
   Runbutton.Enabled = True
   chkPrintAuto.Enabled = True
   chkHtmlAuto.Enabled = True
   chkListAuto.Enabled = True
   Check1.Enabled = True
   chkAutoRounding.Enabled = True
   autonum% = 0
   If Check1.Value = vbUnchecked Then
      Text2.Enabled = False
      optVis.Enabled = False
      optMishor.Enabled = False
      optAst.Enabled = False
      Label4.Enabled = False
   ElseIf Check1.Value = vbChecked Then
      Text2.Enabled = True
      optVis.Enabled = True
      optMishor.Enabled = True
      optAst.Enabled = True
      Label4.Enabled = True
      End If
      
   If chkPrintAuto.Value Then
      autprint = True
   ElseIf chkHtmlAuto.Value Or chkListAuto.Value Then
      autosave = True
      End If
      
   If Check1.Value = vbChecked Then
      Text2.Enabled = True
      'check for stored pagenum
      If Dir(drivjk$ & "numdirec.txt") <> sEmpty Then
         numd% = FreeFile
         Open drivjk$ & "numdirec.txt" For Input As #numd%
         Input #numd%, newpagenum%
         Text2.Text = newpagenum% + 1
         Close #numd%
         End If
      End If
      
   If Combo1.ListCount < 1 Then
        For i% = RefHebYear% To 6000
           Combo1.AddItem (Trim$(Str$(i%)))
        Next i% '<<<<<<<<<<<<<<<<manage years>>>>>>>>>>>>>>

        hebcal = True
        mydate$ = Date
        lenmydate% = Len(mydate$)  '****fix a YK2 bug!****
        If lenmydate% = 8 Then
           yeartab% = Mid$(mydate$, lenmydate% - 1, 2)
           sthebyr% = yeartab% - 97 + 5758
        ElseIf lenmydate% >= 9 Then
           yeartab% = Mid$(mydate$, lenmydate% - 3, 4)
           sthebyr% = yeartab% - 1997 + 5758
           End If
        Combo1.ListIndex = sthebyr%
        End If
        
   Label1.Enabled = False
   Dir1.Enabled = False
   Drive1.Enabled = False
   Label2.Enabled = False
   Text1.Enabled = False
   AutoCancelbut.Enabled = True
   Option1.Enabled = True
   Option2.Enabled = True
   Option3.Enabled = True
   frmTableType.Enabled = True
   optSunrise.Enabled = True
   optSunset.Enabled = True
   optBoth.Enabled = True
End Sub

Private Sub AutoCancelbut_Click()
   If stage% = 0 Then
     Label3.Enabled = False
     Combo1.Enabled = False
     Runbutton.Enabled = False
     Label1.Enabled = True
     Dir1.Enabled = True
     Drive1.Enabled = True
     Label2.Enabled = True
     Text1.Enabled = True
     AutoCancelbut.Enabled = False
     Label4.Enabled = False
     Check1.Enabled = False
     Text2.Enabled = False
     optVis.Enabled = False
     optMishor.Enabled = False
     optAst.Enabled = False
     Label3.Enabled = False
     Option1.Enabled = False
     Option2.Enabled = False
     Option3.Enabled = False
     Text2.Text = numautolst%
     frmTableType.Enabled = False
     optSunrise.Enabled = False
     optSunset.Enabled = False
     optBoth.Enabled = False
  Else
     autocancel = True
     End If
End Sub

Private Sub Check1_Click()
   If Check1.Value = vbUnchecked Then
      Text2.Enabled = False
      Label4.Enabled = False
   ElseIf Check1.Value = vbChecked Then
      Text2.Enabled = True
      Label4.Enabled = True
      End If
End Sub

Private Sub Astroncheck_Click()
   If Astroncheck.Value = vbChecked Then
      astronplace = True
      AstronForm.Visible = True
      'ret = SetWindowPos(AstronForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = vbChecked Then
      calnode.Visible = True
      End If
End Sub

Private Sub chkAutoRounding_Click()
  If chkAutoRounding.Value = vbChecked Then
     cmbRounding.Enabled = True
     cmbRounding.Clear
     cmbRounding.AddItem "1"
     cmbRounding.AddItem "5"
     cmbRounding.AddItem "6"
     cmbRounding.AddItem "10"
     cmbRounding.AddItem "15"
     cmbRounding.AddItem "30"
     cmbRounding.AddItem "60"
     cmbRounding.ListIndex = 0
  Else
     cmbRounding.Enabled = False
     End If
End Sub

Private Sub chkHtmlAuto_Click()
  chkPrintAuto.Value = False
  chkListAuto.Value = False
  Check1.Visible = True
  txtCivil.Visible = False
  PDFprinter = False
End Sub

Private Sub chkListAuto_Click()
  chkPrintAuto.Value = False
  chkHtmlAuto.Value = False
  Check1.Visible = False
  txtCivil.Visible = True
  PDFprinter = False
End Sub



Private Sub chkObs_Click()
   If chkObs.Value = vbChecked Then
'      Call SunriseSunset.chkObst_Click
      AddObsTime = 1
   Else
      AddObsTime = 0
      End If
End Sub

Private Sub chkPdfPrinter_Click()
   chkHtmlAuto.Value = False
   chkListAuto.Value = False
   Check1.Visible = True
   txtCivil.Visible = False
   autoprint = True
   PDFprinter = True
   chkOnePage.Visible = True
End Sub

Private Sub chkPrintAuto_Click()
   chkHtmlAuto.Value = False
   chkListAuto.Value = False
   Check1.Visible = True
   txtCivil.Visible = False
   autoprint = True
   PDFprinter = False
End Sub


Private Sub dir1_click()
      Text1.Text = Dir1.List(Dir1.ListIndex)
      If Not calnearsearchVis Then eros = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Drive1_Change
' Author    : Dr-John-K-Hall
' Date      : 6/1/2017
' Purpose   : Changes directory
'---------------------------------------------------------------------------------------
'
Private Sub Drive1_Change()
   On Error GoTo Drive1_Change_Error

   Dir1.Path = Drive1.Drive    ' When drive changes, set directory path.
   ChDir Drive1.Drive + "\"
   Dir1.ListIndex = 0
   Text1.Text = Dir1.List(Dir1.ListIndex)
  On Error GoTo popmsgbox
10:  'If LTrim$(RTrim$(Drive1.Drive)) <> currentdrive Or errorfnd And initdir = False Then
     '  errorfnd = False
     '  currentdrive = LTrim$(RTrim$(Drive1.Drive))
     '  List1.Clear
'    '  Display the names in A:\ that represent directories.
     '  mypath = currentdrive
     '  myname = Dir(mypath, vbDirectory)   ' Retrieve the first entry.
     '  Do While myname <> sEmpty   ' Start the loop.
     '    ' Ignore the current directory and the encompassing directory.
     '    If myname <> "." And myname <> ".." Then
     '    ' Use bitwise comparison to make sure MyName is a directory.
     '    If (GetAttr(mypath & myname) And vbDirectory) = vbDirectory Then
     '         List1.AddItem myname  ' Display entry only if it
     '         End If  ' it represents a directory.
     '    End If
     '    myname = Dir    ' Get next entry.
     '  Loop
     '  Text1.Text = List1.Text
      Text1.Text = Dir1.List(Dir1.ListIndex)
       GoTo 100
       'End If
popmsgbox:
  MsgBox "Error encountered while trying to open file, please retry.", _
     vbExclamation, "Cal Program"
     errorfnd = True
     Close
     GoTo 10
100:
initdir = False

   On Error GoTo 0
   Exit Sub

Drive1_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Drive1_Change of Form Caldirectories"
End Sub

Private Sub ExitButton_Click()
   On Error GoTo ExitButton_Click_Error

   Screen.MousePointer = vbDefault
   Close
   If Katz = True Then GoTo 55
   On Error GoTo exer50
   'For i% = 0 To Forms.Count - 1
   '   Unload Forms(i%)
   'Next i%
   'Unload Caldirectories
   'Set Caldirectories = Nothing      'clear memory
exer50:
 'now erase tables in fordtm
  'fordtm\netz and/or fordtm\skiy are empty
  If Dir(drivfordtm$ + "netz\*.*") <> sEmpty Then Kill drivfordtm$ + "netz\*.*"
  If Dir(drivfordtm$ + "skiy\*.*") <> sEmpty Then Kill drivfordtm$ + "skiy\*.*"
  If Dir(drivcities$ + "ast\netz\*.*") <> sEmpty Then Kill drivcities$ + "ast\netz\*.*"
  If Dir(drivcities$ + "ast\skiy\*.*") <> sEmpty Then Kill drivcities$ + "ast\skiy\*.*"
  If Dir(drivjk$ + "netzskiy.*") <> sEmpty Then Kill drivjk$ + "netzskiy.*"
  If Katz = True Then
     Caldirectories.Visible = False
     AstronForm.Visible = True
     If Katz = True And katztotal% > 0 And katztotal% <= AstronForm.Combo1.ListCount - 1 Then
        AstronForm.Combo1.ListIndex = katztotal%
        AstronForm.Text1.Text = astcoord(1, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text2.Text = astcoord(2, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text3.Text = astcoord(3, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text4.Text = astrplaces$(AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text5.Text = astcoord(4, AstronForm.Combo1.ListIndex + 1)
        If astcoord(5, AstronForm.Combo1.ListIndex + 1) = 0 Then
           AstronForm.Option2.Value = True
        ElseIf astcoord(5, AstronForm.Combo1.ListIndex + 1) = 1 Then
           AstronForm.Option1.Value = True
           End If
     ElseIf Katz = True And katztotal% > AstronForm.Combo1.ListCount - 1 Then
        katztotal% = 0
        AstronForm.Combo1.ListIndex = AstronForm.Combo1.ListCount - 1
        AstronForm.Text1.Text = astcoord(1, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text2.Text = astcoord(2, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text3.Text = astcoord(3, AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text4.Text = astrplaces$(AstronForm.Combo1.ListIndex + 1)
        AstronForm.Text5.Text = astcoord(4, AstronForm.Combo1.ListIndex + 1)
        If astcoord(5, AstronForm.Combo1.ListIndex + 1) = 0 Then
           AstronForm.Option2.Value = True
        ElseIf astcoord(5, AstronForm.Combo1.ListIndex + 1) = 1 Then
           AstronForm.Option1.Value = True
           End If
        End If
     Exit Sub
  Else
     CalMDIform.Visible = True
     Caldirectories.Visible = False
     Exit Sub
     End If
     
'*****************end the program end*****************
55  Close
    myfile = Dir(drivfordtm$ + "busy.cal")
    If myfile <> sEmpty Then Kill drivfordtm$ + "busy.cal"
         
    lognum% = FreeFile
    Open drivjk$ + "calprog.log" For Append As #lognum%
    If Err.Number <> 0 Then
       Print #lognum%, "Encountered error #: " & Trim$(Str$(Err.Number))
       Print #lognum%, Err.Description
       End If
    Print #lognum%, "Program termination called from Caldirectories:Exitbutton"
    Close #lognum%
         
    For i% = 0 To Forms.Count - 1
      Unload Forms(i%)
    Next i%
          
    'kill timer
    If lngTimerID <> 0 Then lngTimerID = KillTimer(0, lngTimerID)

    'end program abruptly
    End

   On Error GoTo 0
   Exit Sub

ExitButton_Click_Error:
    If Err.Number = 70 Then
       Resume Next
       End If
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExitButton_Click of Form Caldirectories"

End Sub

Private Sub Form_Load()
   'version: 04/08/2003
   
   If Not calnodevis Then eros = False

    Screen.MousePointer = vbHourglass
    mydate = Date
    mont = Val(Mid$(mydate, 1, 2))
    posit% = InStr(1, mydate, "/")
    dayy = Val(Mid$(mydate, posit% + 1, 2))
    posit% = InStr(posit% + 1, mydate, "/")
    yrr = Val(Mid$(mydate, posit% + 1, Len(mydate) - posit%))
    If Len(mydate) - posit% = 4 Then
    ElseIf Len(mydate) - posit% = 2 Then
       If yrr = 99 Then
          yrr = 1900 + yrr
       Else
          yrr = 2000 + yrr
          End If
       End If
        
    'If yrr < 2020 Then yrr = 2000 + Val(Mid$(mydate, posit% + 1, len(mydate) - posit%))
    newday = yrr + mont / 12 + dayy / (12 * 30)

'**********************changes************
GoTo cd500
'*****************************************

   frmSplash.Show
'***************************PASSWORD*********************************
   ' check for password override file c:\jk\pwr.ovr--if it exists in any form
   ' then don't need password. If it doesn't exist then request password
   ' for years 2010 and latter.
   ' FOR requested UPDATES, provide a pwr.ovr file.<----------
'*********************************************************************
   If Dir(drivjk$ + "pwr.ovr") = sEmpty And newday >= 2005 + 1 / (12 * 30) Then '1999 + 10 / 12 + 11 / (12 * 30) Then '+ 8 / 12 + 13 / (12 * 30) Then
       frmSplash.Label1.Visible = False
       frmSplash.NewLabel.Visible = True
       frmSplash.NewLabel.Caption = Chr$(34) + "לוחות " + Chr$(34) + "חי"
       title$ = "לוח " + Chr$(34) + "חי" + Chr$(34)
       address$ = "© לוחות חי, טל/פקס: 5713765(02).  גרסא: " & Str(datavernum) & "." & Str(progvernum)
       If optionheb = False Then
          address$ = "© Luchos Chai, Fahtal  56, Jerusalem  97430; Tel/Fax: +972-2-5713765. Version: " & Str(progvernum) & "." & Str(datavernum)
          End If
       Calyoufm.Show 1
       If goahead = False Then
         ExitButton_Click
         End If
    Else
       title$ = "לוח " + Chr$(34) + "בכורי יוסף" + Chr$(34)
       'address$ = "מדרש בכורי יוסף, ירושלים"
       address$ = "מדרש בכורי יוסף, ת.ד. 35078, ירושלים 91350"
       End If
     
   
cd500:
   magnify = False
   'now read in city files
   filcity% = FreeFile
   myfile = Dir(drivcities$ + "citynams_w1255.txt")
   If myfile = sEmpty Then
      response = MsgBox("Caldirectories can't find the list of Israeli cities.  " + _
               "This means that you don't have the option to calculate the VISUAL sunrises and sunsets " + _
               "for these cities.", vbInformation + vbOKCancel, "Cal Program")
      If response = vbCancel Then
         'MsgBox "Caldirectories can't find the city list file: citynams.txt!  ABORTING program...Sorry", vbCritical + vbOKOnly, "Cal Program"
         Caldirectories.ExitButton.Value = True
      Else
         Dir1.Enabled = False
         Drive1.Enabled = False
         Label1.Caption = "Check the box above"
         Autobut.Enabled = False
         Astroncheck.TabIndex = 0
         Astroncheck.ForeColor = QBColor(13)
         numcities% = 0: numheb% = 0
         End If
   Else
      Open drivcities$ + "citynams_w1255.txt" For Input As #filcity%
      numcit% = 0: numcities% = 0: numheb% = 0
      Do Until EOF(filcity%)
         numcit% = numcit% + 1
         Line Input #filcity%, docline$
         If numcit% Mod 2 = 0 Then
            'newhebcalfm.Combo1.AddItem title$ + " לנץ החמה ב" + docline$
            'newhebcalfm.Combo6.AddItem title$ + " לשקיעת החמה ב" + docline$
            numheb% = numheb% + 1
            'remove the "_" from the names
            For inam% = 1 To Len(docline$)
               If Mid$(docline$, inam%, 1) = "_" Then
                  Mid$(docline$, inam%, 1) = " "
                  End If
            Next inam%
            cityhebnames$(numheb%) = docline$
         Else
            numcities% = numcities% + 1
            citynames$(numcities%) = docline$
            End If
      Loop
      Close #filcity%
      End If
   filnetz3% = FreeFile
   Text1.Text = sEmpty
   errorfnd = False
 On Error GoTo popmsgbox
10: If internet = True Then
       Caldirectories.Visible = False
       If eros = True Then
          foundvantage = False
          calnearsearchfm.Visible = True
          calnearsearchfm.Text1 = Str(eroslongitude)
          calnearsearchfm.Text2 = Str(eroslatitude)
          calnearsearchfm.Text3 = searchradius
          calnearsearchfm.Command3.Value = True
          'wait a bit for the program to find the contributing
          'vantage points within the search radius
          Do Until foundvantage = True
             DoEvents
          Loop
          'now wait a bit more
          waitime = Timer + 0.1
          Do Until Timer > waitime
             DoEvents
          Loop
          calnearsearchfm.Command1.Value = 1 'calculate this table
          Exit Sub
          End If
   
       lognum% = FreeFile
       Open drivjk$ + "calprog.log" For Append As #lognum%
       Print #lognum%, "Step #3: Activating SunriseSunset Form "
       Close #lognum%
          
       SunriseSunset.Visible = True
       SunriseSunset.Combo1.Text = Str(yrheb%)
       SunriseSunset.OKbut0.Value = True
       Exit Sub
       End If
    ChDrive defdriv$
    Drive1.Drive = defdriv$ + ":\"
    Dir1.Path = defdriv$ + ":\"
    ChDir defdriv$ + ":\cities"
    currentdrive = Trim$(Drive1.Drive)
    Dir1.ListIndex = 0
    BringWindowToTop (Caldirectories.hwnd)
    'Dir1.List (0)

'' Display the names in c:\ that represent directories.
'  MyPath = currentdrive '"a:\"  ' Set the path.
'  MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
'  Do While MyName <> sEmpty   ' Start the loop.
'    ' Ignore the current directory and the encompassing directory.
'    If MyName <> "." And MyName <> ".." Then
'        ' Use bitwise comparison to make sure MyName is a directory.
'
'       If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
'           List1.AddItem MyName  ' Display entry only if it
'           End If  ' it represents a directory.
'       End If
'    MyName = Dir    ' Get next entry.
'   Loop
   Text1.Text = Dir1.List(Dir1.ListIndex)
GoTo 100
popmsgbox:
  'Caldirectories.Visible = True
'  Unload frmSplash
'  Set frmSplash = Nothing
  If Err.Number >= 68 And Err.Number <= 71 Then
     MsgBox "Caldirectories encountered Disk/Path error " + CStr(Err.Number) + "...try again!", vbExclamation, "Cal Program"
  Else
     MsgBox "Caldirectories encountered unexpected error " + CStr(Err.Number) + "...please retry!", _
     vbExclamation, "Cal Program"
     End If
   Close
   GoTo 10
100:
   Caldirectories.Visible = True
   Label3.Enabled = False
   Combo1.Enabled = False
   
   With Caldirectories
      .Label4.Enabled = False
      .Check1.Enabled = False
      .Text2.Enabled = False
      .optVis.Enabled = False
      .optMishor.Enabled = False
      .optAst.Enabled = False
   End With
   
   If Not calnodevis Then eros = False
   If internet = False Then
      nearcolor = False
      If SunriseSunset.Check3.Value = vbChecked And AddObsTime = 0 Then  'if adding additional time for near obstructions, don't print them in green
         nearcolor = True
      Else
         nearski = False
         nearnez = False
         nearcolor = False
         End If
      End If
   automatic = False
   runningscan = False
   stage% = 0
   autocancel = False
   CN4netz$ = sEmpty
   CN4skiy$ = sEmpty
   nearauto = False
   nearautoedited = False
   Runbutton.Enabled = False
   AutoCancelbut.Enabled = False
   frmSplash.Visible = Fase
   Unload frmSplash
   Set frmSplash = Nothing
   'ret = SetWindowPos(Caldirectories.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   Screen.MousePointer = vbDefault
  ' initdir = True
End Sub


'Private Sub List1_Click()
'   Text1.Text = List1.Text
'End Sub


Private Sub OKbutton_Click()
   Screen.MousePointer = vbHourglass
   startedscan = False
   OKbutton.Enabled = False
   ExitButton.Enabled = False
   Label1.Enabled = False
   Drive1.Enabled = False
   Dir1.Enabled = False
   'List1.Enabled = False
   Text1.Enabled = False
   SunriseSunset.Timer1.Enabled = False
   SunriseSunset.Show
   SunriseSunset.Enabled = True
   SunriseSunset.ProgressBar1.Enabled = True
   SunriseSunset.ProgressBar1.Visible = True
   'SunriseSunset.Label3.Visible = False
   currentdir = Trim$(Text1.Text)
   If InStr(currentdir, "eros") <> 0 Then
      eros = True
      geo = True
      End If
   SunriseSunset.Visible = True
   SunriseSunset.Check1.Enabled = True
   SunriseSunset.Check2.Enabled = True
   SunriseSunset.Check3.Enabled = True
   SunriseSunset.Check4.Enabled = True
   SunriseSunset.Check5.Enabled = True
   SunriseSunset.Combo1.Enabled = True
   SunriseSunset.Label2.Enabled = True
   SunriseSunset.Cancelbut.Enabled = True
   If SunriseSunset.Label1.Caption <> sEmpty Then SunriseSunset.Label1.Caption = captmp$
   SunriseSunset.Label1.Enabled = True
   SunriseSunset.OKbut0.Enabled = True
   SunriseSunset.ProgressBar1.Visible = False
   SunriseSunset.Option1.Enabled = True
   SunriseSunset.Option2.Enabled = True
   SunriseSunset.Label2.Enabled = True
   SunriseSunset.Label4.Enabled = True
   SunriseSunset.Label5.Enabled = True
   SunriseSunset.Label6.Enabled = True
   SunriseSunset.Label7.Enabled = True
   SunriseSunset.Text1.Enabled = True
   SunriseSunset.UpDown1.Enabled = True
   SunriseSunset.Text2.Enabled = True
   SunriseSunset.UpDown2.Enabled = True
   SunriseSunset.Option3.Enabled = True
   SunriseSunset.Option4.Enabled = True
   If optionheb = False Then
      SunriseSunset.Option4.Value = True
      End If
      
   If hebcal Then
      Option1b = True
   Else
      Option2b = True
      End If
      
   If Option1b = True Then
      SunriseSunset.Option1.Value = True
      End If
   If Option2b = True Then
      SunriseSunset.Option2.Value = True
      End If
   'SunriseSunset.Combo1.Text = 5758
   'ret = SetWindowPos(SunriseSunset.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   If suntop% <> 0 Then SunriseSunset.Top = suntop%
   Screen.MousePointer = vbDefault
End Sub


Private Sub optAst_Click()
   visauto = False
   mishorauto = False
   astauto = True
   chkObs.Enabled = False
   chkObs.Value = vbUnchecked
End Sub

Private Sub optBoth_Click()
  SunriseCalc = True
  SunsetCalc = True
  With SunriseSunset
     .Check1.Value = vbChecked
     .Check2.Value = vbChecked
  End With
End Sub

Private Sub optEnglish_Click()
   'output civil calendars as output of automatic operation
   Combo1.Clear
   'english calendar
   For i% = 1600 To 2235
      Combo1.AddItem (Trim$(Str$(i%)))
   Next i% '<<<<<<<<<<<<<<<<manage years>>>>>>>>>>>>>>

   hebcal = False
   mydate$ = Date
   Combo1.ListIndex = Year(Date) - 1600
End Sub

Private Sub optHebrew_Click()
   'hebrew calendars as output for automatic operation
   Combo1.Clear
   For i% = RefHebYear% To 6000
      Combo1.AddItem (Trim$(Str$(i%)))
   Next i% '<<<<<<<<<<<<<<<<manage years>>>>>>>>>>>>>>

   hebcal = True
   mydate$ = Date
   lenmydate% = Len(mydate$)  '****fix a YK2 bug!****
   If lenmydate% = 8 Then
      yeartab% = Mid$(mydate$, lenmydate% - 1, 2)
      sthebyr% = yeartab% - 97 + 5758
   ElseIf lenmydate% >= 9 Then
      yeartab% = Mid$(mydate$, lenmydate% - 3, 4)
      sthebyr% = yeartab% - 1997 + 5758
      End If
   Combo1.ListIndex = sthebyr%

End Sub

Private Sub Option1_Click()
   nearauto = False
   nearautoedited = False
End Sub

Private Sub Option2_Click()
'   On Error GoTo operror
   autoNoCDcheck = False
   nearauto = True
   nearautoedited = False
   If autosave Then
      autoNoCDcheck = True
      Exit Sub
      End If
   'check for CD containing prom directory if using the CD
   autoNoCDcheck = False
   response = MsgBox("Are the prom/prof directories on a CD?", vbQuestion + vbYesNo + vbDefaultButton2, "Cal Program")
   If response = vbNo Then
      autoNoCDcheck = True
      Exit Sub
      End If
55  defa$ = Mid$("cdefghijklmnop", Caldirectories.Drive1.ListCount - 1, 1)
    If automatic = True Then
       cddrivlet$ = defa$
       GoTo 57
       End If
    myvalue = InputBox("What is the CD-ROM's drive letter?", "Cal Program", defa$, 3000, 900)
    cddrivlet$ = myvalue
    If cddrivlet$ = sEmpty Then
       Option1.Value = True
       Exit Sub
       End If
    MsgBox "Make sure that disk containing the PROM directory is loaded in the CD-ROM drive. BEWARE that this search might detect virtual sunset obstructions produced by glitches present in some near-seashore cities' .004 files!", vbInformation + vbOKOnly, "Cal Program"
57  chkh = Dir(cddrivlet$ + ":\prom", vbDirectory)
    If chkh = sEmpty Then
       response = MsgBox("PROM directory not found, please make sure that CD-ROM containing the PROM directory is loaded properly in the CD-ROM.", vbCritical + vbOKCancel, "Cal Program")
       If response = vbOK Then
          GoTo 55
       Else
          Option1.Value = True
          End If
       End If

'   response = MsgBox("Make sure the CD containing the prom directory is loaded into the CD drive", vbInformation + vbOKCancel, "Cal Program")
'   If response = vbCancel Then
'      Option1.Value = True
'      End If
'op210:   myfile = Dir("h:\prom", vbDirectory)
'   Exit Sub
'operror:
'   response = MsgBox("Can't find the prom directory!, Make sure the CD containing the prom directory is loaded into the CD drive. Try again?", vbOKCancel + vbCritical, "Cal Program")
'   If response = vbCancel Then
'      Option1.Value = True
'      Exit Sub
'      End If
'  Resume
End Sub

Private Sub Option3_Click()
   autoNoCDcheck = False
   nearauto = True
   'use edited list
   nearautoedited = True
   If autosave Then
      autoNoCDcheck = True
      Exit Sub
      End If
End Sub

Private Sub optMishor_Click()
   visauto = False
   mishorauto = True
   astauto = False
   chkObs.Enabled = False
   chkObs.Value = vbUnchecked
End Sub

Private Sub optSunrise_Click()
  SunriseCalc = True
  SunsetCalc = False
  With SunriseSunset
     .Check1.Value = vbChecked
     .Check2.Value = vbUnchecked
  End With
End Sub

Private Sub optSunset_Click()
  SunriseCalc = False
  SunsetCalc = True
  With SunriseSunset
     .Check1.Value = vbUnchecked
     .Check2.Value = vbChecked
  End With
End Sub

Private Sub optVis_Click()
   visauto = True
   mishorauto = False
   astauto = False
   chkObs.Enabled = True
   chkObs.Value = vbChecked
End Sub

Private Sub Runbutton_Click()
   If Not SunriseCalc And Not SunsetCalc Then
      Call MsgBox("Click on the sunrise, sunset, or both option buttons to define which type of calculation to do", vbInformation, "Calculation Type")
      Exit Sub
      End If
      
   If runningscan = True Then Exit Sub
   If autocancel = True Then GoTo 950
   yrheb% = Val(Combo1.Text)
   
   If chkListAuto.Value = vbChecked And Not BeginCivilRun Then
      If Trim$(txtCivil) = sEmpty Then
         'make sure that folder has a year identifier
         MsgBox "Enter a suffix in the text box provided, e.g., ''2012-2015''", vbOKOnly + vbInformation, "Chai Program"
         Exit Sub
      Else
         'determine beginning year and number of years
         BeginningYear$ = Mid$(txtCivil.Text, 1, 4)
         EndYear$ = Mid$(txtCivil.Text, Len(txtCivil.Text) - 3, 4)
         NumCivilYears% = Val(EndYear$) - Val(BeginningYear$) + 1
         If NumCivilYears% < 0 Then
            MsgBox "The format of the suffix in the text box is incorrect." _
                   & vbCrLf & "It should be like ''2012-2015''.", vbOKOnly + vbInformation, App.title
            Exit Sub
         Else
            Select Case MsgBox("Please check the following values:" _
                               & vbCrLf & "" _
                               & vbCrLf & "Beginning year:  " & BeginningYear$ _
                               & vbCrLf & "End year:  " & EndYear$ _
                               & vbCrLf & "Number of years:  " & Str$(NumCivilYears%) _
                               & vbCrLf & "" _
                               & vbCrLf & "Is this correct?" _
                               , vbYesNo Or vbInformation Or vbDefaultButton1, App.title)
            
                Case vbYes
                   NumCivilYearsInc% = 0
                   BeginCivilRun = True
                Case vbNo
                   BeginCivilRun = False
                   Exit Sub
            End Select
            End If
         End If
      
      End If
    
   If chkHtmlAuto.Value = vbChecked Or chkListAuto.Value = vbChecked Or PDFprinter Then GoTo 920 'don't use citynams.lst
   
   myfile = Dir(drivcities$ + "citynams_w1255.lst")
   If myfile = sEmpty Then
      MsgBox "Can't find list of cities--ABORT the automatic mode!", vbExclamation + vbOKOnly, "Cal Program"
      Label3.Enabled = False
      Combo1.Enabled = False
      Runbutton.Enabled = False
   Else
      automatic = True
      If stage% = 0 Then MsgBox "It is advised that you close down all other applications, if any, before continuing.", vbInformation + vbOKOnly, "Cal Program"
      'read list and start doing cities
      GoSub autorun
      End If
   Exit Sub
autorun:
   If stage% = 1 Then GoTo 900
   Label1.Enabled = False
   fillst1% = FreeFile
   For i% = 0 To Dir1.ListCount - 1
      Dir1.ListIndex = i%
      If Dir1.List(Dir1.ListIndex) + "\" = drivcities$ Then '+ ":\cities" Then
         Dir1.Path = Dir1.List(Dir1.ListIndex)
         Dir1.ListIndex = -1
         Exit For
         End If
   Next i%
   numautocity% = 0 'determine how many automatic files to do
   Open drivcities$ + "citynams_w1255.lst" For Input As #fillst1%
   Do Until EOF(fillst1%)
      Line Input #fillst1%, doclin$
      numautocity% = numautocity% + 1
   Loop
   Close #fillst1%
   numautolst% = 0
   newpagenum% = 0
   If Text2.Text <> "1" And chkOnePage.Value = vbUnchecked Then
      response = MsgBox("Do you want to start from city number: " + Text2.Text + " ?", vbYesNo + vbQuestion, "Cal Program")
      If response = vbYes Then
         newpagenum% = Val(Text2.Text) - 1
         numautolst% = Val(Text2.Text) - 1
         End If
  ElseIf chkOnePage.Value = vbChecked Then
      newpagenum% = Val(Text2.Text) - 1
      numautolst% = Val(Text2.Text) - 1
      autocancel = False
      End If
      
900 fillst1% = FreeFile
    Open drivcities$ + "citynams_w1255.lst" For Input As #fillst1%
    If numautolst% + 1 > numautocity% Then GoTo 950
    'If EOF(fillst1%) Then GoTo 999
      Dir1.ListIndex = 0
      For i% = 1 To numautolst% + 1
        Input #fillst1%, hebnam$, currentdir, tblmesag%, s1blk, s2blk
      Next i%
      Close #fillst1%
      numautolst% = numautolst% + 1
      For i% = 0 To Dir1.ListCount - 1
         If Dir1.List(i%) = currentdir Then
            Dir1.ListIndex = i%
            Exit For
            End If
      Next i%
      'newer% = 1
      newer$ = currentdir
      calTime = Timer + 2#
      Do While calTime > Timer
         DoEvents
      Loop
      runningscan = True
      Caldirectories.OKbutton.Value = True
      If nearauto = False Then
         SunriseSunset.OKbut0.Value = True
      ElseIf nearauto = True Then
         SunriseSunset.Check3.Value = vbChecked
         SunriseSunset.OKbut0.Value = True
         End If
      stage% = 1
      
      If chkPrintAuto.Value = vbChecked Then GoTo 999
920:  'make html directories if required
      
      If chkHtmlAuto.Value = vbChecked Then
         htmldir$ = "zemanim_"
         If mishorauto Then
            htmldir$ = htmldir$ & "mishor_"
         ElseIf astauto Then
            htmldir$ = htmldir$ & "ast_"
         Else 'use default
            htmldir$ = htmldir$ & "visible_"
            End If
            
         htmldir$ = drivjk$ & "html_city_tables\" & htmldir$ & Trim$(Combo1.Text)
         
         mydir$ = Dir(htmldir$, vbDirectory)
         If mydir$ = sEmpty Then 'make directory
            MkDir htmldir$
            End If
            
         End If
         
      If chkListAuto.Value = vbChecked Then
      
         htmldir$ = "zemanim_"
         If mishorauto Then
            htmldir$ = htmldir$ & "mishor_"
         ElseIf astauto Then
            htmldir$ = htmldir$ & "ast_"
         Else 'use default
            htmldir$ = htmldir$ & "visible_"
            End If
            
         htmldir$ = drivjk$ & "html_city_tables\" & htmldir$ & txtCivil.Text
         
         mydir$ = Dir(htmldir$, vbDirectory)
         If mydir$ = sEmpty Then 'make directory
            MkDir htmldir$
            End If
      
         End If
         
      If chkPdfPrinter.Value = vbChecked Then
         dircreate$ = drivjk$ & "\pdf_city_tables\" & Trim$(Caldirectories.Combo1.Text)
         If Dir(dircreate$, vbDirectory) = sEmpty Then
            MkDir drivjk$ & "\pdf_city_tables\" & Trim$(Caldirectories.Combo1.Text)
            End If
         End If
         
      'open citysorthebrew.txt and process all the files listed there
      myfile = Dir(drivcities & "citysortheb_w1255.txt")
      myfile2 = Dir(drivcities & "citynams_w1255.txt")
      If myfile = sEmpty Or myfile2 = sEmpty Then
         MsgBox ("Can't find citynams_w1255.txt, or citysortheb.txt, aborting automatic mode...")
         Label3.Enabled = False
         Combo1.Enabled = False
         Runbutton.Enabled = False
      Else
         GoSub autorun2
         Exit Sub
         End If

autorun2:
      If stage% = 1 Then GoTo 930
      Label1.Enabled = False
      For i% = 0 To Dir1.ListCount - 1
         Dir1.ListIndex = i%
         If Dir1.List(Dir1.ListIndex) + "\" = drivcities$ Then '+ ":\cities" Then
            Dir1.Path = Dir1.List(Dir1.ListIndex)
            Dir1.ListIndex = -1
            Exit For
            End If
      Next i%
      numautocity% = 0 'determine how many automatic files to do
      fillst1% = FreeFile
      Open drivcities$ + "citysortheb_w1255.txt" For Input As #fillst1%
      Do Until EOF(fillst1%)
         Line Input #fillst1%, doclin$
         numautocity% = numautocity% + 1
      Loop
      Close #fillst1%
      numautolst% = 0
      newpagenum% = 0
      If Text2.Text <> "1" And chkOnePage.Value = vbUnchecked Then
         response = MsgBox("Do you want to start from city number: " + Text2.Text + " ?", vbYesNo + vbQuestion, "Cal Program")
         If response = vbYes Then
            newpagenum% = Val(Text2.Text) - 1
            numautolst% = Val(Text2.Text) - 1
            End If
      ElseIf chkOnePage.Value = vbChecked Then
         newpagenum% = Val(Text2.Text) - 1
         numautolst% = Val(Text2.Text) - 1
         autocancel = False
         End If
      
930:  If numautolst% + 1 > numautocity% Then GoTo 950
      filcit% = FreeFile
      If nearautoedited Then
        'only calculate tables for the edited city list
        Open drivcities & "citysorthebedited_w1255.txt" For Input As #filcit%
        For i% = 1 To numautolst% + 1
           Input #filcit%, doclin$
        Next i%
        Close #filcit%
      Else
        'calculate tables for all the cities
        Open drivcities & "citysortheb_w1255.txt" For Input As #filcit%
        For i% = 1 To numautolst% + 1
           Input #filcit%, doclin$
        Next i%
        Close #filcit%
        End If
        
      numautolst% = numautolst% + 1
         
      'now open citynams_w1255.txt and find equivalent English name
      filnam% = FreeFile
      Open drivcities & "citynams_w1255.txt" For Input As #filnam%
      Dir1.ListIndex = 0
       Do Until EOF(filnam%)
          Input #filnam%, cityAutoEng$
          Input #filnam%, cityAutoHeb$
          If cityAutoHeb$ = doclin$ Then
             Close #filnam%
             found% = 1
             Exit Do
             End If
       Loop
       If found% = 1 Then
          currentdir = cityAutoEng$
       Else
          Close #filnam%
          MsgBox ("Can't find equivalent English directory name of :" & Trim$(doclin$) & "!  Aborting...")
          GoTo 950
          End If
       
       'postiion directory listing at currentdir
       For i% = 0 To Dir1.ListCount - 1
          If Dir1.List(i%) = drivcities & currentdir Then
            Dir1.ListIndex = i%
            Exit For
            End If
      Next i%
      newer$ = currentdir
      calTime = Timer + 2#
      Do While calTime > Timer
         DoEvents
      Loop
      runningscan = True
      automatic = True
      Caldirectories.OKbutton.Value = True
      If nearauto = False Then
         AddObsTime = 0
         SunriseSunset.OKbut0.Value = True
      ElseIf nearauto = True Then
         If chkObs.Value = vbChecked Then
            AddObsTime = 1
            SunriseSunset.Check3.Value = vbUnchecked
            SunriseSunset.chkObst.Value = vbChecked
            SunriseSunset.OKbut0.Value = True
         Else
            AddObsTime = 0
            SunriseSunset.Check3.Value = vbChecked
            SunriseSunset.chkObst.Value = vbUnchecked
            SunriseSunset.OKbut0.Value = True
            End If
          End If
      stage% = 1
      If chkOnePage.Value Then autocancel = True
      GoTo 999
  
'     If numautolst% + 1 > numautocity% Then

950:
      If chkListAuto.Value = vbChecked And Not autocancel Then
         If NumCivilYearsInc% + 1 < NumCivilYears% Then
            'begin next year
            NumCivilYearsInc% = NumCivilYearsInc% + 1
            Combo1.ListIndex = Combo1.ListIndex + 1
            yrheb% = Val(Combo1.Text)
            SunriseSunset.Combo1.Text = Caldirectories.Combo1.Text
            numautolst% = 0
            stage% = 0
            GoTo 920
         Else
            BeginCivilRun = False
            'beep for a while to warn of completion
            waitime = Timer
            Do Until Timer > waitime + 60
               Beep
               DoEvents
            Loop
            End If
         End If
         
      automatic = False

      If autosave Then 'write closing to TOC

           fnum% = FreeFile
           
           myfile = Dir(drivjk$ & "html_city_tables\" & "TOC.html")
           If myfile <> sEmpty Then
              Open drivjk$ & "html_city_tables\" & "TOC.html" For Append As #fnum%
              Print #fnum%, "</BODY>"
              Print #fnum%, "</HTML>"
              Close #fnum%
              End If
              
           autosave = False
           End If
           
      If autoprint Then autoprint = False

      With Caldirectories
      .Text2.Text = CStr(numautolst%) + 1
         .Runbutton.Enabled = False
         .Combo1.Enabled = False
         .AutoCancelbut.Enabled = False
         .Dir1.ListIndex = -1
         .Label4.Enabled = False
         .Check1.Enabled = False
         .Text2.Enabled = False
         .optVis.Enabled = False
         .optMishor.Enabled = False
         .optAst.Enabled = False
         .Label3.Enabled = False
         .Dir1.Enabled = True
         .Drive1.Enabled = True
         .Label2.Enabled = True
         .Text1.Enabled = True
      End With
      
      Option1.Enabled = False
      Option2.Enabled = False
      Option3.Enabled = False
      
      'ret = SetWindowPos(Caldirectories.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      runningscan = False
      autocancel = False
      Close
      stage% = 0
'      End If
'   GoTo 900
999   Label1.Enabled = True
End Sub

Private Sub Text1_Change()
      If astronplace = False Then Text1.Text = Dir1.List(Dir1.ListIndex)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ExitButton_Click
End Sub

