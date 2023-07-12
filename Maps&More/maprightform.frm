VERSION 5.00
Begin VB.Form maprightform 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Places"
   ClientHeight    =   6705
   ClientLeft      =   7815
   ClientTop       =   1845
   ClientWidth     =   3360
   Icon            =   "maprightform.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Load Map Coordinates"
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame frmTrig 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Trig Point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1275
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   3135
      Begin VB.CommandButton cmdTrigUndo 
         Caption         =   "&Undo last Trig Point Correction"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdTrigPoint 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save &Trig point Coordinates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   300
         Width           =   2655
      End
   End
   Begin VB.Frame frmPlace 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Placlist Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   2960
      Width           =   3135
      Begin VB.CommandButton rightPLACEbut 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Add this Entry to Placlist"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   1080
         TabIndex        =   11
         Text            =   "0.5"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "km"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Calculations begin after distance:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton rightEXITbut 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6180
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton rightSavebut 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save Map Point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   2475
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "hgt(m):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ITMy:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ITMx:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "maprightform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLoad_Click()
   'load coordinate text boxes with map coordinates and map name
   Text2 = Maps.Text5
   Text3 = Maps.Text6
   Text4 = Maps.Text7
   If mapfile$ <> sEmpty Then
      For i% = Len(mapfile$) - 4 To 1 Step -1
        CH$ = Mid$(mapfile$, i%, 1)
        If CH$ = "\" Then
           rootname$ = Mid$(mapfile$, i% + 1, Len(mapfile$) - 4 - i%)
           Exit For
           End If
      Next i%
      Text1 = rootname$
      End If
   
End Sub

Private Sub cmdTrigPoint_Click()
   'record trig point
   kmxTrig = Text2
   kmyTrig = Text3
   hgtTrig = Text4
   Call Form_QueryUnload(0, 0)
End Sub

Private Sub cmdTrigUndo_Click()
   'undo last trig point fix
   If Dir(israeldtm & ":\dtm\" & CHMNEO & "_" & Month(Date) & Day(Date) & Year(Date)) <> sEmpty Then
      Close
      FileCopy israeldtm & ":\dtm\" & CHMNEO & "_" & Month(Date) & Day(Date) & Year(Date), israeldtm & ":\dtm\" & CHMNEO
      CHMNEO = sEmpty
   Else
mrf100: response = MsgBox("Backup file for today not found." & vbLf & _
             "Do you want to input name of backup file?", _
             vbExclamation + vbYesNoCancel, "Maps&More")
      If response = vbYes Then
         res = InputBox("Input name of backup file", "Backup File Name", _
                         israeldtm & ":\dtm\" & CHMNEO & "_" & Month(Date) & Day(Date) & Year(Date), 6250)
         If res <> sEmpty Then
            If Dir(res) <> sEmpty Then
               Close
               FileCopy res, israeldtm & ":\dtm\" & CHMNEO
               CHMNEO = sEmpty
            Else
               GoTo mrf100
               End If
            End If
         End If
      End If
   Call Form_QueryUnload(0, 0)
End Sub

Private Sub Form_Load()
   'enable buttons and frames
   If Maps.mnuTrigDrag.Checked Then
      'disenable place frame
      frmPlace.Visible = False
      Text1.Visible = False
      lblName.Visible = False
      lblName.Enabled = False
      rightSavebut.Visible = False
      maprightform.Height = 4130 + 380
      frmTrig.Visible = True
      frmTrig.Top = 1700 + 380
      rightEXITbut.Top = 3140 + 380
      maprightform.Caption = "Define Trig Point"
   Else
      'disenable trig pont frame
      frmTrig.Visible = False
      frmPlace.Visible = True
      maprightform.Height = 5580 + 380
      frmPlace.Top = 2580 + 380
      rightEXITbut.Top = 4580 + 380
      maprightform.Caption = "Add places"
      End If
      
End Sub

Private Sub rightEXITbut_Click()
    Call Form_QueryUnload(i%, j%)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lmag As Long
    Unload maprightform
    Set maprightform = Nothing
    If magbox = True Then 'restore it to top of z order
'      ret = SetWindowPos(mapMAGfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'      ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      BringWindowToTop (mapMAGfm.hWnd)
      BringWindowToTop (mapPictureform.hWnd)
       'lmag = FindWindow(vbNullString, mapMAGfm.Caption)
       'ret = BringWindowToTop(lmag)
    Else
      If world = False Then
         ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         lResult = FindWindow(vbNullString, terranam$)
         If lResult > 0 Then
'             ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             BringWindowToTop (lResult)
             End If
      Else
'         ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         BringWindowToTop (mapPictureform.hWnd)
         lResult = FindWindow(vbNullString, "3D Viewer")
         If lResult > 0 Then
'             ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             BringWindowToTop (lResult)
             End If
         End If
       End If
End Sub

Private Sub rightPLACEbut_Click()
   If world = True And Val(Text4.Text) = 0 Then
      response = MsgBox("You must enter a non-zero height for the chosen place!", vbExclamation + vbOKOnly, "Maps&More")
      Exit Sub
      End If
   Screen.MousePointer = vbHourglass
   plac$ = drivjk_c$ + "placlist.txt"
   myfile = Dir(plac$)
   filplac% = FreeFile
   If myfile = sEmpty Then
      Open plac$ For Output As #filplac%
   Else
      Open plac$ For Append As #filplac%
      End If
   ln1% = Len(LTrim(RTrim$(Text1.Text)))
   If ln1% >= 20 Then
      txt1$ = "'" + Mid$(LTrim$(RTrim$(Text1.Text)), 1, 20) + "'"
   Else
      txt1$ = "'" + LTrim$(RTrim$(Text1.Text)) + String(20 - ln1%, " ") + "'"
      End If
   If Text5.Text = 0.5 Then
      txt3$ = "0"
   Else
      txt3$ = Text5.Text
      End If
   If world = False Then
      txt1$ = txt1$ + "," + LTrim$(RTrim$(Text2.Text * 0.001)) + "," + _
      LTrim$(RTrim$((Text3.Text - 1000000) * 0.001)) + "," + _
      LTrim$(RTrim$(Text4.Text)) + ",1,0," + txt3$
   Else
      txt1$ = txt1$ + "," + LTrim$(RTrim$(Text2.Text)) + "," + _
      LTrim$(RTrim$(Text3.Text)) + "," + _
      LTrim$(RTrim$(Text4.Text)) + ",1,0," + txt3$
      End If
   
   Print #filplac%, txt1$
   Close #filplac%
   Screen.MousePointer = vbDefault
End Sub


Private Sub rightSavebut_Click()
   'save this entry in the SKYLIGHT.sav file
   'first check if there is such a file, and if there is append to it
   'and if not, then open a new file
   Screen.MousePointer = vbHourglass
   filsav% = FreeFile
   If world = False Then
      filsav1$ = drivjk_c$ + "SkyLight.sav"
   Else
      filsav1$ = drivjk_c$ + "skyworld.sav"
      End If
   myfile = Dir(filsav1$)
   If myfile = sEmpty Then
      Open filsav1$ For Output As #filsav%
   Else
      Open filsav1$ For Append As #filsav%
      End If
   Write #filsav%, Text1.Text, Val(Text2.Text), Val(Text3.Text), Val(Text4.Text)
   Close #filsav%
   Screen.MousePointer = vbDefault
End Sub
