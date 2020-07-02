VERSION 5.00
Begin VB.Form mapsplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4185
   ClientLeft      =   2325
   ClientTop       =   2175
   ClientWidth     =   7335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "mapsplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7305
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6480
         TabIndex        =   7
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "topo maps are copyrighted property of the Survey of Isreal"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5640
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "©"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "©"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   4
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CNK 1999-2001"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6060
         TabIndex        =   3
         Top             =   3630
         Width           =   1215
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   2
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Maps && More"
         BeginProperty Font 
            Name            =   "David"
            Size            =   48
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   840
         TabIndex        =   1
         Top             =   1440
         Width           =   6135
      End
      Begin VB.Image Image1 
         Height          =   4245
         Left            =   0
         Picture         =   "mapsplash.frx":000C
         Top             =   0
         Width           =   8715
      End
   End
End
Attribute VB_Name = "mapsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub form_load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Label3.Caption = "CNK 1999-" + LTrim$(Str(Year(Date)))
   ' lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

