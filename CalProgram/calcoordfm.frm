VERSION 5.00
Begin VB.Form calcoordfm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input your coordinates"
   ClientHeight    =   3615
   ClientLeft      =   3885
   ClientTop       =   3765
   ClientWidth     =   4650
   Icon            =   "calcoordfm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   180
      TabIndex        =   2
      Top             =   60
      Width           =   4215
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1620
         TabIndex        =   9
         ToolTipText     =   "name of the place (English)"
         Top             =   1860
         Width           =   2235
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1620
         TabIndex        =   7
         Text            =   "0"
         ToolTipText     =   "height in meters"
         Top             =   1320
         Width           =   2235
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1620
         TabIndex        =   5
         Text            =   "0"
         ToolTipText     =   "(degrees latitude; positive for northern hemisphere)"
         Top             =   780
         Width           =   2235
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1620
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "(degrees longitude; positive for western hemisphere)"
         Top             =   300
         Width           =   2235
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "height:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "latitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "longitude:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   4
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2820
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2820
      Width           =   1155
   End
End
Attribute VB_Name = "calcoordfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    eroscitylong(citynodenum%) = Val(Text4.Text)
    eroscitylat(citynodenum%) = Val(Text5.Text)
    eroscityhgt(citynodenum%) = Val(Text6.Text)
    eroscity$ = Text1.Text
    eroscityarea(citynodenum%) = eroscityarea(citynodenum% - 1)
    erosareabat = eroscityarea(citynodenum%)
    eroslongitude = eroscitylong(citynodenum%)
    eroslatitude = eroscitylat(citynodenum%)
    eroscountry$ = eroscountries(citynodenum% - 1)
    calnode.StatusBar1.Panels(2) = "User Input" & ". lon: " & eroscitylong(citynodenum%) & ", lat: " & eroscitylat(citynodenum%) & ", hgt: " & eroscityhgt(citynodenum%) & " in city area: " & eroscityarea(citynodenum% - 1)
    Call Form_QueryUnload(i%, j%)
    calnearsearchfm.Visible = True
    BringWindowToTop (calnearsearchfm.hwnd)
    calnearsearchfm.StatusBar1.Visible = False
    calnearsearchfm.Text1 = eroscitylong(citynodenum%)
    calnearsearchfm.Text2 = eroscitylat(citynodenum%)
    calnearsearchfm.Text3 = 8
End Sub

Private Sub Command2_Click()
   Call Form_QueryUnload(i%, j%)
End Sub

Private Sub Form_Load()
   'version: 04/08/2003

    Text4.Text = Str$(eroscitylong(citynodenum%))
    Text5.Text = Str$(eroscitylat(citynodenum%))
    Text6.Text = Str$(eroscityhgt(citynodenum%))
    Text1.Text = eroscity$
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload calcoordfm
End Sub

