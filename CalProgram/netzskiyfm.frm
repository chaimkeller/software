VERSION 5.00
Begin VB.Form netzskiyfm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   5640
   ClientLeft      =   2940
   ClientTop       =   1995
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton NetzSkiyOkbut0 
      BackColor       =   &H0080FF80&
      Caption         =   "&Accept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      Picture         =   "netzskiyfm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear &All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton NetzskiyCancelbut 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   5040
      Picture         =   "netzskiyfm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
   Begin VB.ListBox Netzskiylist 
      Height          =   4110
      Index           =   2
      ItemData        =   "netzskiyfm.frx":0884
      Left            =   120
      List            =   "netzskiyfm.frx":088B
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "netzskiyfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAll_Click()
   For i% = 6 To nn4%
      Netzskiylist(2).Selected(i% - 6) = True
   Next i%
End Sub

Private Sub cmdClear_Click()
   For i% = 6 To nn4%
      Netzskiylist(2).Selected(i% - 6) = False
   Next i%
End Sub

Private Sub NetzskiyCancelbut_Click(Index As Integer)
   'version: 04/08/2003
  
'  exit from rountine and go back to main menu
   Screen.MousePointer = vbDefault
   netzskiyok = False
   Unload netzskiyfm
   Set netzskiyfm = Nothing
   SunriseSunset.Visible = False
   'If calnode.Visible = True Then
   If eroscityflag = True Then
      Caldirectories.Visible = False
      Exit Sub
      End If
   Caldirectories.Label1.Enabled = True
   Caldirectories.Drive1.Enabled = True
   Caldirectories.Dir1.Enabled = True
   'Caldirectories.List1.Enabled = True
   Caldirectories.Text1.Enabled = True
   Caldirectories.OKbutton.Enabled = True
   Caldirectories.ExitButton.Enabled = True
   'ret = SetWindowPos(Caldirectories.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   'myfile = Dir(drivjk$+":\jk\netzskiy.tm3")
   'If myfile <> sEmpty Then
   '   Kill drivjk$+":\jk\netzskiy.tm3"
   '   End If
   'myfile = Dir(drivjk$+":\jk\netzskiy.tm4")
   'If myfile <> sEmpty Then
   '   Kill drivjk$+":\jk\netzskiy.tm4"
   '   End If
End Sub

Private Sub NetzskiyOKbut0_Click()
   Screen.MousePointer = vbHourglass
   If automatic = True Then
      waittime = Timer + 1#
      Do While waittime > Timer
         DoEvents
      Loop
      End If
   numchecked% = 0
   For i% = 6 To nn4%
      If Netzskiylist(2).Selected(i% - 6) = True Then
         nchecked%(i% - 5) = 1
         numchecked% = numchecked% + 1
      Else
         nchecked%(i% - 5) = 0
         End If
   Next i%
   netzskiyok = True
   Unload netzskiyfm
   Set netzskiyfm = Nothing
End Sub


