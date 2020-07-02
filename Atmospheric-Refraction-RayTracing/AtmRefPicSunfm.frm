VERSION 5.00
Begin VB.Form AtmRefPicSunfm 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   5385
   Begin VB.PictureBox picRef 
      AutoRedraw      =   -1  'True
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "AtmRefPicSunfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   With AtmRefPicSunfm
    picRef.Left = .Left + 10
    picRef.Width = .Width - 20
    picRef.TOP = 10
    picRef.Height = .Height - 20
   End With
End Sub

Private Sub Form_Resize()
   With AtmRefPicSunfm
    picRef.Left = .Left + 10
    picRef.Width = .Width - 20
    picRef.TOP = 10
    picRef.Height = .Height - 20
   End With
End Sub
