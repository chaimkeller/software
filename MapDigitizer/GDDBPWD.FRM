VERSION 5.00
Begin VB.Form GDfrmDBPWD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Required"
   ClientHeight    =   1260
   ClientLeft      =   7395
   ClientTop       =   4920
   ClientWidth     =   2820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2016133
   Icon            =   "GDDBPWD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   1260
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   795
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   795
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   315
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Enter database password:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1905
   End
End
Attribute VB_Name = "GDfrmDBPWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Password Required"
Const BUTTON1 = "&OK"
Const BUTTON2 = "&Cancel"
Const Label1 = "&Enter database password:"
'>>>>>>>>>>>>>>>>>>>>>>>>

Public PWD As String

Private Sub cmdCancel_Click()
  PWD = vbNullString
  PwdCancel = True
  'Hide
End Sub

Private Sub Form_Load()
  Me.Caption = FORMCAPTION
  cmdOK.Caption = BUTTON1
  cmdCancel.Caption = BUTTON2
  lblLabels(0).Caption = Label1
End Sub

Private Sub cmdOK_Click()
  PWD = txtPassword.Text
  'Hide
End Sub

Private Sub txtPassword_GotFocus()
  txtPassword.SelStart = 0
  txtPassword.SelLength = Len(txtPassword.Text)
End Sub
