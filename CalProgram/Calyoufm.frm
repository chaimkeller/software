VERSION 5.00
Begin VB.Form Calyoufm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cal's Password"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Calyoufm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   40
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Hint: have you forgotten הכרת הטוב?"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Password and a carriage return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Calyoufm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chnum%(20), numstroke%, passmode%
Private Sub Form_Load()
   'version: 04/08/2003
  
   Screen.MousePointer = vbDefault
   goahead = False
   passmode% = 0 '0 = feelings OK but need update
                 '1 = feelings hurt
   If passmode% = 0 Then
      Text1.ToolTipText = "Hint: You need to update!, call 02-5713765"
      End If
End Sub
Private Sub text1_keypress(KeyAscii As Integer)
   'record keys inputed, and look for carriage return
   If KeyAscii = 13 Then GoTo 50
   'convert to * character
   lent% = Len(Text1)
   Text1 = String(lent%, "*")
   numstroke% = numstroke% + 1
   If numstroke% > 20 Then
      goahead = False
      Beep
      numstroke% = 0
      Text1.Text = sEmpty
      Exit Sub
      End If
   chnum%(numstroke%) = KeyAscii
   Exit Sub
   
50: If passmode% = 1 Then
     If numstroke% = 4 Then
       'check if the short password is correct
       If (chnum%(1) = 110 Or chnum%(1) = 78) And _
          (chnum%(2) = 101 Or chnum%(2) = 69) And _
          (chnum%(3) = 97 Or chnum%(3) = 65) And _
          (chnum%(4) = 108 Or chnum%(4) = 76) Then
          goahead = True
          Call Form_QueryUnload(i%, j%)
          Exit Sub
       Else
          goahead = False
          Beep
          numstroke% = 0
          Text1 = sEmpty
          response = MsgBox("Have you forgotten הכרת הטוב?", vbInformation + vbOKOnly, "Cal's Password Hint")
          'Call form_queryunload(i%, j%)
          Exit Sub
          End If
    ElseIf numstroke% = 13 Then  'check for longer password
       If (chnum%(1) = 104 Or chnum%(1) = 72) And _
          (chnum%(2) = 97 Or chnum%(2) = 65) And _
          (chnum%(3) = 107 Or chnum%(3) = 75) And _
          (chnum%(4) = 97 Or chnum%(4) = 65) And _
          (chnum%(5) = 114 Or chnum%(5) = 82) And _
          (chnum%(6) = 111 Or chnum%(6) = 79) And _
          (chnum%(7) = 115 Or chnum%(7) = 83) And _
          (chnum%(8) = 32 Or chnum%(8) = 32) And _
          (chnum%(9) = 104 Or chnum%(9) = 72) And _
          (chnum%(10) = 97 Or chnum%(10) = 65) And _
          (chnum%(11) = 116 Or chnum%(11) = 84) And _
          (chnum%(12) = 111 Or chnum%(12) = 79) And _
          (chnum%(13) = 118 Or chnum%(13) = 86) Then
          goahead = True
          Call Form_QueryUnload(i%, j%)
          Exit Sub
       Else
          goahead = False
          Beep
          numstroke% = 0
          Text1 = sEmpty
          response = MsgBox("Have you forgotten הכרת הטוב?", vbInformation + vbOKOnly, "Cal's Password Hint")
          'Call form_queryunload(i%, j%)
          Exit Sub
          End If
    ElseIf numstroke% = 9 Then  'check for short password
       If chnum%(1) = 228 And _
          chnum%(2) = 235 And _
          chnum%(3) = 248 And _
          chnum%(4) = 250 And _
          chnum%(5) = 32 And _
          chnum%(6) = 228 And _
          chnum%(7) = 232 And _
          chnum%(8) = 229 And _
          chnum%(9) = 225 Then
          goahead = True
          Call Form_QueryUnload(i%, j%)
          Exit Sub
       Else
          goahead = False
          Beep
          numstroke% = 0
          Text1 = sEmpty
          response = MsgBox("Have you forgotten הכרת הטוב?", vbInformation + vbOKOnly, "Cal's Password Hint")
          'Call form_queryunload(i%, j%)
          Exit Sub
          End If
    Else
       goahead = False
       Beep
       numstroke% = 0
       Text1 = sEmpty
       response = MsgBox("Have you forgotten הכרת הטוב?", vbInformation + vbOKOnly, "Cal's Password Hint")
       'Call form_queryunload(i%, j%)
       Exit Sub
       End If
 ElseIf passmode% = 0 Then
    If numstroke% = 4 Then
       'check if the short password is correct
       If (chnum%(1) = 110 Or chnum%(1) = 78) And _
          (chnum%(2) = 101 Or chnum%(2) = 69) And _
          (chnum%(3) = 97 Or chnum%(3) = 65) And _
          (chnum%(4) = 108 Or chnum%(4) = 76) Then
          goahead = True
          Call Form_QueryUnload(i%, j%)
          Exit Sub
       Else
          goahead = False
          Beep
          numstroke% = 0
          Text1 = sEmpty
          response = MsgBox("THE DATA BASE USED BY THIS PROGRAM MUST BE UPDATED!, call 02-5713765 for further inforamtion.", vbRetryCancel + vbCritical, "Cal's Password Message")
          If response = vbCancel Then
             Call Form_QueryUnload(i%, j%)
             End If
          Exit Sub
          End If
   Else
      goahead = False
      Beep
      numstroke% = 0
      Text1 = sEmpty
      response = MsgBox("THE DATA BASE USED BY THIS PROGRAM MUST BE UPDATED!, call 02-5713765 for further inforamtion.", vbRetryCancel + vbCritical, "Cal's Password Message")
      If response = vbCancel Then
         Call Form_QueryUnload(i%, j%)
         End If
       Exit Sub
       End If
     End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set Calyoufm = Nothing
End Sub


