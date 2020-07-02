VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form calnode 
   Caption         =   "Country and citys' list (calculate visible sunrise times using the GTOPO10 DTM)"
   ClientHeight    =   3900
   ClientLeft      =   2505
   ClientTop       =   2835
   ClientWidth     =   7155
   Icon            =   "calnode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   7155
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   3585
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   7276
            MinWidth        =   2469
            Text            =   "Choose your home city, or input your coordinates"
            TextSave        =   "Choose your home city, or input your coordinates"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDB 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5106
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4140
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":0442
            Key             =   "metropolitisold"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":075C
            Key             =   "city"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":0A76
            Key             =   "stateold"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":0D90
            Key             =   "state"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":10AA
            Key             =   "metro"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":13C4
            Key             =   "USAimage"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":16DE
            Key             =   "Canadaimage"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":19F8
            Key             =   "Englandimage"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":1D12
            Key             =   "Franceimage"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":202C
            Key             =   "Italyimage"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":2346
            Key             =   "Russiaimage"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":2660
            Key             =   "USAcountry"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":297A
            Key             =   "Canadacountry"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":2C94
            Key             =   "Francecountry"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":2FAE
            Key             =   "Italycountry"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":32C8
            Key             =   "Englandcountry"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":3B1A
            Key             =   "Mexicoimage"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":3E34
            Key             =   "Mexicocountry"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":414E
            Key             =   "Switzerlandimage"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":4468
            Key             =   "Belgiumimage"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":4782
            Key             =   "Denmarkimage"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":4A9C
            Key             =   "Netherlandsimage"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":4DB6
            Key             =   "Greececountry"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":77B0
            Key             =   "Belgiumcountry"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":A352
            Key             =   "Denmarkcountry"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":D158
            Key             =   "Netherlandscountry"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":FC4A
            Key             =   "Switzerlandcountry"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":11E74
            Key             =   "Greeceimage"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":178CE
            Key             =   "Uruguaycountry"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":D09D8
            Key             =   "Uruguayimage"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":E16A2
            Key             =   "Brazilcountry"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":195424
            Key             =   "Brazilimage"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":1A6C06
            Key             =   "Argentinaimage"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":1DAB78
            Key             =   "Israelimage"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":1DAE92
            Key             =   "Israelcountry"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":1DAFEC
            Key             =   "Austriaimage"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "calnode.frx":1DB43E
            Key             =   "Germanyimage"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "calnode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim placlat As Double, placlon As Double, plachgt As Double
Dim eroslongitude As Double, eroslatitude As Double

'Private erosstatenum As Integer
'Private erosstates(50) As String
'Private erosstatesindex(50) As Integer
'Private eroscitylong(1000) As Single
'Private eroscitylat(1000) As Single
'Private eroscityhgt(1000) As Single
'Private eroscityarea(1000) As String
'Private eroscitynum(50) As Integer
'Private eroscities(50, 50) As String
'Private eroslocatnum(50, 50) As Integer
'Private eroslocat(50, 50, 150) As String '(states,cities,locations)
Private Sub Form_Load()
   'version: 01/28/2004
      
   Israelflag% = 0 '<--EY
  
    'read the eroscity.sav file to determine states and city areas
    'then read the individual sav files and load the individual locations
    If eroscityflag = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    erosstatenum = -1
    'For i% = 1 To 50
    '   eroscitynum(i%) = -1
    '   For j% = 1 To 50
    '      eroslocatnum(i%, j%) = -1
    '   Next j%
    'Next i%
    ecdir$ = drivcities$ & "eros\"
    myfile = Dir(ecdir$ & "eroscity.sav")
    If myfile = sEmpty Then
       Screen.MousePointer = vbDefault
       response = MsgBox("Can't find the eroscity.sav file", vbCritical + vbOKOnly, "Maps & More")
       Unload calnode
       Exit Sub
    Else
    '****The USA******
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "USA"
       country$ = "USA"
       namesize% = Len(country$)
       mNode.Image = "USAimage"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
       End If
900:
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Canada"
       country$ = "Canada"
       namesize% = Len(country$)
       mNode.Image = "Canadaimage"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
    
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "England"
       country$ = "England"
       namesize% = Len(country$)
       mNode.Image = "Englandimage"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
    
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "France"
       country$ = "France"
       namesize% = Len(country$)
       mNode.Image = "Franceimage"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
    
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Italy"
       country$ = "Italy"
       namesize% = Len(country$)
       mNode.Image = "Italyimage"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize

       'tvwDB.Sorted = True
       'ecfilnum% = FreeFile
       'Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       'tvwDB.LabelEdit = False
       'Set mNode = tvwDB.Nodes.Add()
       'mNode.Sorted = True
       'tvwDB.LabelEdit = False
       'mNode.Text = "Russia"
       'country$ = "Russia"
       'namesize% = Len(country$)
       'mNode.Image = "Russiaimage"
       'countryindex = mNode.Index
       'erosstatenum = 0
       'GoSub organize
    
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Mexico"
       country$ = "Mexico"
       namesize% = Len(country$)
       mNode.Image = "Mexicoimage"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
    
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Switzerland"
       country$ = "Switzerland"
       namesize% = Len(country$)
       mNode.Image = "Switzerlandimage"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize

       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Belgium"
       country$ = "Belgium"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize

       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Denmark"
       country$ = "Denmark"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize

       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Netherlands"
       country$ = "Netherlands"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
       
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Greece"
       country$ = "Greece"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize

       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Uruguay"
       country$ = "Uruguay"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
       
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Brazil"
       country$ = "Brazil"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
       
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Argentina"
       country$ = "Argentina"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
       
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Austria"
       country$ = "Austria"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
       
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Germany"
       country$ = "Germany"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       GoSub organize
       
       
      '********Israel must always be last**********
       tvwDB.Sorted = True
       ecfilnum% = FreeFile
       Open ecdir$ & "eroscity.sav" For Input As #ecfilnum
       tvwDB.LabelEdit = False
       Set mNode = tvwDB.Nodes.Add()
       mNode.Sorted = True
       tvwDB.LabelEdit = False
       mNode.Text = "Israel"
       country$ = "Israel"
       namesize% = Len(country$)
       mNode.Image = country$ + "image"
       countryindex = mNode.Index
       erosstatenum = 0
       Israelflag% = 1 '<--EY
       GoSub organize
       Israelflag% = 0


950:  Screen.MousePointer = vbDefault
Exit Sub

organize:
       Do Until EOF(ecfilnum%)
          Line Input #ecfilnum%, ecnam$
          myfile = Dir(ecdir$ & ecnam$ & ".sav")
          If myfile <> sEmpty Then
             ec2filnum% = FreeFile
             Open ecdir$ & ecnam$ & ".sav" For Input As #ec2filnum%
          Else
             response = MsgBox("Can't find " & ecnam$ & "'s .sav file!", vbCritical + vbOKOnly, "Maps & More")
             Close
             Exit Sub
             End If
          'parse eroscity name into city area, state, country
          If Mid$(ecnam$, Len(ecnam$) - namesize%, namesize% + 1) = "_" & country$ Then
             nch% = 1
             Do
               cha$ = Mid$(ecnam$, Len(ecnam$) - namesize% - nch%, 1)
               If cha$ = "_" Then
                  Exit Do
               Else
                  nch% = nch% + 1
                  End If
             Loop
             statenam$ = Mid$(ecnam$, Len(ecnam$) - nch% - namesize% + 1, nch% - 1)
             'check if this state is new
             'If statenam$ = "PA" Then
             '   cc = 1
             '   End If
             flag% = -1
             For i% = 0 To erosstatenum
                If statenam$ = erosstates(i%) Then
                   flag% = 0
                   Exit For
                   End If
             Next i%
             If flag% = -1 Then
                erosstatenum = erosstatenum + 1
                erosstates(erosstatenum) = statenam$
                'Set mNode = tvwDB.Nodes.Add()'   //old
                '// new
                'Set mNode = tvwDB.Nodes.Add(countryindex, tvwChild, , statenam$, "state")
                On Error GoTo stateerror
                Set mNode = tvwDB.Nodes.Add(countryindex, tvwChild, , statenam$, country$ + "country")
                '// new
                mNode.Sorted = True
                tvwDB.LabelEdit = False
                'mNode.Text = statenam$'  //old
                'mNode.Image = "state"'   //old
                stateindex = mNode.Index
                erosstatesindex(erosstatenum) = stateindex
             Else
                'erosstatenum = i%
                stateindex = erosstatesindex(i%)
                End If
             'now add the city area name
             intIndex = mNode.Index
             cityarea$ = Mid$(ecnam$, 1, Len(ecnam$) - Len(statenam$) - namesize% - 2)
             Set mNode = tvwDB.Nodes.Add(stateindex, tvwChild, , cityarea$, "metro")
             mNode.Sorted = True
             'nch2% = 1
             'Do
             '  cha$ = Mid$(ecnam$, Len(ecnam$) - 3 - nch% - nch2%, 1)
             '  If cha$ = "_" Then
             '     Exit Do
             '  Else
             '     nch2% = nch2% + 1
             '     End If
             'Loop
             'cityarea$ = Mid$(ecnam$, Len(ecnam$) - nch% - 3 - nch2%, nch2%)
             'eroscitynum(erosstatenum) = eroscitynum(erosstatenum) + 1
             'eroscities(erosstatenum, eroscitynum(erosstatenum)) = cityarea$
             intIndex = mNode.Index
             Do Until EOF(ec2filnum%)
                Input #ec2filnum%, placnam$, placlat, placlon, plachgt
                If country$ = "Israel" Then
                   placnam$ = Replace(placnam$, "_", " ")
                   End If
                'eroslocatnum(erosstatenum, eroscitynum(erosstatenum)) = eroslocatnum(erosstatenum, eroscitynum(erosstatenum)) + 1
                'eroslocat(erosstatenum, eroscitynum(erosstatenum), eroslocatnum(erosstatenum, eroscitynum(erosstatenum))) = placnam$
                If InStr(eroscountry$, "Israel") <> 0 Then
                   If placlat > 1000 Then placlat = (placlat - 1000000) * 0.001
                   If placlon > 1000 Then placlon = placlon * 0.001
                   End If
                Set mNode = tvwDB.Nodes.Add(intIndex, tvwChild)
                mNode.Text = placnam$
                mNode.Image = "city"     ' Image from ImageList.
                citynodenum% = mNode.Index
                eroscitylong(citynodenum%) = placlon
                eroscitylat(citynodenum%) = placlat
                eroscityhgt(citynodenum%) = plachgt
                eroscityarea(citynodenum%) = ecnam$
                eroscountries(citynodenum%) = country$
             Loop
             Set mNode = tvwDB.Nodes.Add(intIndex, tvwChild)
             mNode.Text = "***User Inputed Coordinates***"
             mNode.Image = "city"     ' Image from ImageList.
             End If
       Loop
       Close
Return

stateerror:
     'couldn't find the bitmap of the country in the image list, so use the default one
     Set mNode = tvwDB.Nodes.Add(countryindex, tvwChild, , statenam$, "state")
     Resume Next
     

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Unload Me
   Set calnode = Nothing
   If CalMDIform.Visible = False Then
      CalMDIform.Visible = True
      End If
End Sub

Private Sub Form_Resize()
   tvwDB.Width = calnode.Width - 30
   tvwDB.Height = calnode.Height - 30 - StatusBar1.Height
End Sub


Private Sub tvwDB_Collapse(ByVal Node As MSComCtlLib.Node)
   eroscity$ = sEmpty
End Sub

Private Sub tvwDB_DblClick()
   If userinput = False And InStr(eroscity$, "_") = 0 And Len(eroscity$) > 2 And eroscity$ <> sEmpty Then
        Select Case eroscity$
           Case "USA", "England", "Canada", "France", "Italy", "Russia"
              Exit Sub
           Case Else
        End Select
        
        calcoordfm.Visible = False
        calnearsearchfm.Visible = True
        BringWindowToTop (calnearsearchfm.hwnd)
        calnearsearchfm.StatusBar1.Visible = False
        calnearsearchfm.Text1 = eroslongitude
        calnearsearchfm.Text2 = eroslatitude
        calnearsearchfm.Text3 = 8 'cities outside Israel
        If eroscountry$ = "Israel" Then calnearsearchfm.Text3 = 1 'Israel neighborhoods
        End If
End Sub

Private Sub tvwDB_NodeClick(ByVal Node As Node)
    'StatusBar1.Panels(2) = "Index = " & Node.Index & " Text:" & Node.Text
Select Case Node.Text
   Case "USA", "Canada", "England", "France", "Italy", "Russia"
      Exit Sub
   Case "AZ", "CA", "CO", "CT", "FL", "GA", "IL", "IN", "MA", "MD", "MI", "MN", "MO", "NY", "NJ", "OH", "PA", "RI", "TN", "TX", "WA", "WI"
      Exit Sub
   Case Else
End Select

If Node.Text = "***User Inputed Coordinates***" Then
       citynodenum% = Node.Index
       calcoordfm.Visible = True
       If eroscity$ = sEmpty Then eroscity$ = sEmpty 'Node.Text
       ret = SetWindowPos(calcoordfm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
       erosareabat = eroscityarea(Node.Index)
       userinput = True
       Exit Sub
       End If
    
StatusBar1.Panels(2) = Node.Text & ". lon: " & eroscitylong(Node.Index) & ", lat: " & eroscitylat(Node.Index) & ", hgt: " & eroscityhgt(Node.Index) & " in city area: " & eroscityarea(Node.Index)
eroslongitude = eroscitylong(Node.Index)
eroslatitude = eroscitylat(Node.Index)
erosareabat = eroscityarea(Node.Index)
eroscountry$ = eroscountries(Node.Index)
eroscity$ = Node.Text
userinput = False

'calnode.Caption = Node.Text
End Sub
    
