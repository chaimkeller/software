VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6975
   ScaleWidth      =   4770
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_MouseMove(Button As Integer, _
   Shift As Integer, X As Single, Y As Single)
   'Form2.Text1 = "kmx = " + Str$(X)
   'Form2.Text2 = "kmy = " + Str$(Y)
   'Form2.Text3 = "hgt ="
   'Form2.Text4 = Time
End Sub


