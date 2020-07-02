VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3225
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7320
      Begin VB.Label NewLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "לוחות חי"
         BeginProperty Font 
            Name            =   "David"
            Size            =   27.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00C0C000&
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "CNK 2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   1875
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "לזמני הנץ והשקיעה הנראים"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   675
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "לוח בכורי יוסף"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   5055
      End
      Begin VB.Image imgLogo 
         Height          =   3225
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7335
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Set frmSplash = Nothing
End Sub

Private Sub Form_Load()
   'version: 04/08/2003
  
    'lblCopyright.Caption = "Internet Version: 1.0.0"
    'lblCopyright.Caption = Year(Date)
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyright.Caption = "CNK " & Trim$(Year(Now))
    'lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
    Set frmSplash = Nothing
End Sub

