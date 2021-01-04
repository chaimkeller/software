VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{825967DA-1756-11D3-B695-ED78B587442C}#15.0#0"; "FlexListBox.ocx"
Begin VB.Form Form3 
   Caption         =   "TestForm"
   ClientHeight    =   5076
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5916
   LinkTopic       =   "Form3"
   ScaleHeight     =   5076
   ScaleWidth      =   5916
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Seek ListItem   [ENTER to Select]  [Arrow DOWN]  [Arrow UP]"
      Height          =   636
      Left            =   144
      TabIndex        =   16
      Top             =   96
      Width           =   5676
      Begin VB.CheckBox chkSelectOnMatch 
         Caption         =   "SelectOnMatch"
         Height          =   192
         Left            =   3984
         TabIndex        =   19
         Top             =   336
         Width           =   1452
      End
      Begin VB.CheckBox chkExactMatch 
         Caption         =   "Exact Match"
         Height          =   252
         Left            =   2784
         TabIndex        =   18
         Top             =   288
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   96
         TabIndex        =   17
         Top             =   192
         Width           =   2604
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ListItems under Cursor Position"
      Height          =   1644
      Left            =   3168
      TabIndex        =   4
      Top             =   3360
      Width           =   2652
      Begin VB.Label Label1 
         Caption         =   "TooltipText"
         Height          =   156
         Index           =   4
         Left            =   96
         TabIndex        =   14
         Top             =   1344
         Width           =   1116
      End
      Begin VB.Label Label1 
         Caption         =   "Selected"
         Height          =   156
         Index           =   3
         Left            =   96
         TabIndex        =   13
         Top             =   1092
         Width           =   1116
      End
      Begin VB.Label Label1 
         Caption         =   "Enabeled"
         Height          =   156
         Index           =   2
         Left            =   96
         TabIndex        =   12
         Top             =   828
         Width           =   1116
      End
      Begin VB.Label Label1 
         Caption         =   "Value"
         Height          =   156
         Index           =   1
         Left            =   96
         TabIndex        =   11
         Top             =   576
         Width           =   1116
      End
      Begin VB.Label Label1 
         Caption         =   "Text"
         Height          =   156
         Index           =   0
         Left            =   96
         TabIndex        =   10
         Top             =   324
         Width           =   1116
      End
      Begin VB.Label lblItems 
         Height          =   156
         Index           =   4
         Left            =   1344
         TabIndex        =   9
         Top             =   1356
         Width           =   1260
      End
      Begin VB.Label lblItems 
         Height          =   156
         Index           =   3
         Left            =   1344
         TabIndex        =   8
         Top             =   1104
         Width           =   1260
      End
      Begin VB.Label lblItems 
         Height          =   156
         Index           =   2
         Left            =   1344
         TabIndex        =   7
         Top             =   852
         Width           =   1260
      End
      Begin VB.Label lblItems 
         Height          =   156
         Index           =   1
         Left            =   1344
         TabIndex        =   6
         Top             =   588
         Width           =   1260
      End
      Begin VB.Label lblItems 
         Height          =   156
         Index           =   0
         Left            =   1344
         TabIndex        =   5
         Top             =   336
         Width           =   1260
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Selected Items"
      Height          =   348
      Left            =   192
      TabIndex        =   3
      Top             =   3312
      Width           =   2892
   End
   Begin FlexList.FlexListBox FlexListBox2 
      Height          =   1308
      Left            =   192
      TabIndex        =   2
      Top             =   3696
      Width           =   2892
      _ExtentX        =   5101
      _ExtentY        =   2307
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListStyle       =   2
      Locked          =   -1  'True
   End
   Begin FlexList.FlexListBox FlexListBox1 
      Height          =   2076
      Left            =   144
      TabIndex        =   1
      Top             =   768
      Width           =   5724
      _ExtentX        =   10097
      _ExtentY        =   3662
      ForeColorUnselected=   16776960
      BackColorUnselected=   13533278
      ForeColorSelected=   16711680
      BackColorSelected=   16776960
      BeginProperty FontSelected {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   -1  'True
      Columns         =   3
      ListStyle       =   2
      SelectOnMatch   =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(in)Visible Disabeled"
      Height          =   300
      Left            =   3936
      TabIndex        =   0
      Top             =   2928
      Width           =   1932
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4224
      Top             =   3696
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":1150
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":1E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":22A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":26F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":2B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":2F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":33F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":3844
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":3C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":40EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":4540
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":4994
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":4DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":523C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":5690
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   192
      TabIndex        =   15
      Top             =   2928
      Width           =   3372
   End
   Begin VB.Line Line1 
      X1              =   5856
      X2              =   192
      Y1              =   3264
      Y2              =   3264
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Developed by Ted Schopenhouer   ted.schopenhouer@12move.nl

'with ideas and suggestions from Hans Scholten Wonen@Wonen.com
'                                Eric de Decker secr VB Belgie
'                                E.B. our API Killer  ProFinance Woerden
'                                Willem secr VBgroup Ned vbg@vbgroup.nl

'This sources may be used freely without the intention of commercial distribution. For all
'For ALL other use of this control YOU MUST HAVE PERMISSION of the developer.

Private bNoChange As Boolean

Private Sub chkExactMatch_Click()
FlexListBox1.ExactMatch = chkExactMatch.Value And vbChecked
Text1.SetFocus
End Sub

Private Sub chkSelectOnMatch_Click()
FlexListBox1.SelectOnMatch = chkSelectOnMatch.Value And vbChecked
Text1.SetFocus
End Sub

Private Sub Command1_Click()
Dim item As List   'class of the flexListBox

With FlexListBox1

   For Each item In .List
      
      With item
         
         'with the .Visible ListMember you can make ListItems disappear
         'as if they don't exist
         If Not .Enabeled Then .Visible = Not .Visible
      
      End With
   
   Next
   
   .Refresh   'refresh view
   
End With
End Sub

Private Sub Command2_Click()
Dim item As List

FlexListBox2.Clear

For Each item In FlexListBox1.List
   
   With item
      
      If .Selected Then
         FlexListBox2.AddItem .Text, , .Enabeled, .UnderlayingValue, .ToolTipText, .ItemPicture, .DownPicture, .Visible
      End If
      
   End With

Next

FlexListBox2.Refresh

End Sub


Private Sub FlexListBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, Listitem As String)
Call DispListItems(CLng(Listitem))
End Sub

Private Sub FlexListBox1_PopUpItems(MenuItemsArray() As Variant, Listitem As Long)

'Example of the use of an popup menu
'   If listItem = 33 Then MenuItemsArray = array("a","b","c")
'or If SomeThingElse Then MenuItemsArray = ........

If Listitem Mod 2 <> 0 Then
   
   MenuItemsArray = Array("Delete ListItem " & CStr(Listitem), "-", IIf(FlexListBox1.List.item(Listitem).Enabeled, "Disable", "Enable") & " ListItem " & CStr(Listitem), "-", "Hide ListItem " & CStr(Listitem), "AddItem")
                           
                           
ElseIf Listitem < 50 Then
   
   MenuItemsArray = Array("One", "Two", "Three", "Four", "Five")

Else
   
   MenuItemsArray = Array("Developed By Ted ", "With Ideas From Hans", "and API Killer EB", "Also Starring Willem", "And Eric as BetaTester")

End If

End Sub

Private Sub FlexListBox1_PopUpItemsClick(MenuIndex As Integer, Listitem As Long)
Const DELETE = 0
Const DISenABLE = 2
Const HIDE = 4
Const AddItem = 5
'Example of the use of an popup menu
'If listItem = 33 And MenuIndex = DELETE Then Delete ......

If Listitem Mod 2 <> 0 Then

   With FlexListBox1
      
      Select Case MenuIndex
         
         Case DELETE
            .RemoveItem Listitem
            
         Case DISenABLE
            .List.item(Listitem).Enabeled = Not .List.item(Listitem).Enabeled
            
         Case HIDE
            .List.item(Listitem).Visible = False
         
         Case AddItem
            .AddItem CStr(Time)
            
      End Select
         
      .Refresh
      
   End With
   
End If

Label2 = "Example For ListItem " & CStr(Listitem) & " MenuIndex " & CStr(MenuIndex)

End Sub

Private Sub DispListItems(Listitem As Long)
Static lItem As Long
If lItem <> Listitem And Listitem > 0 Then
   lItem = Listitem
   With FlexListBox1.List.item(lItem)
      lblItems(0) = .Text
      lblItems(1) = .UnderlayingValue
      lblItems(2) = .Enabeled
      lblItems(3) = .Selected
      lblItems(4) = .ToolTipText
   End With
End If

End Sub


Private Sub Form_Load()
Dim i       As Integer
Dim i2      As Integer
Dim Tmp     As String

Randomize

For i = 1 To 100
   
   Tmp = ""
   
   For i2 = 1 To 5
      Tmp = Tmp & Chr(65 + Int(Rnd * 24))
   Next
   
   FlexListBox1.AddItem CStr(i) & Tmp, CStr(i), True, Rnd * i, "tooltip " & CStr(i), ImageList1.ListImages(Int(Rnd * 17) + 1).ExtractIcon, ImageList1.ListImages(Int(Rnd * 17) + 1).ExtractIcon
   
   FlexListBox1.List.item(i).Enabeled = Int(Rnd * 24) Mod 7 <> 0
   
Next

Call Form4.ShowFlexValues(FlexListBox1)

Form4.Show
Form3.ZOrder

chkExactMatch.Value = FlexListBox1.ExactMatch And vbChecked
chkSelectOnMatch.Value = FlexListBox1.SelectOnMatch And vbChecked

End Sub

Private Sub Text1_Change()
If Not bNoChange Then FlexListBox1.Text = Text1.Text
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
   FlexListBox1.CurrPosItem = FlexListBox1.CurrPosItem + 1
   Call DispListItems(FlexListBox1.CurrPosItem)
   KeyCode = 0
ElseIf KeyCode = vbKeyUp Then
   FlexListBox1.CurrPosItem = FlexListBox1.CurrPosItem - 1
   Call DispListItems(FlexListBox1.CurrPosItem)
   KeyCode = 0
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   With FlexListBox1
      If .CurrPosItem > 0 Then
         If .ListIndex + 1 = .CurrPosItem Then
            bNoChange = True
            Text1.Text = ""
            .List.item(.CurrPosItem).Selected = False
            .ListIndex = -1
            bNoChange = False
         Else
            Text1.Text = .List.item(.CurrPosItem()).Text
            Text1.SelStart = Len(Text1.Text)
         End If
      End If
   End With
   KeyAscii = 0
End If
End Sub
