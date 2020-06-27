VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplines 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Courbes SPLINES"
   ClientHeight    =   11370
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   17925
   Icon            =   "Splines.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11370
   ScaleWidth      =   17925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameVue 
      Caption         =   "Vue "
      Height          =   1575
      Left            =   16920
      TabIndex        =   34
      Top             =   7680
      Width           =   855
      Begin VB.OptionButton OptionVue 
         Caption         =   "YZ"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   615
      End
      Begin VB.OptionButton OptionVue 
         Caption         =   "XZ"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton OptionVue 
         Caption         =   "XY"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.TextBox txtEditPoint 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Enter pour confirmer"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picQuadrillage 
      BackColor       =   &H00E0E0E0&
      Height          =   10380
      Left            =   255
      MouseIcon       =   "Splines.frx":0442
      MousePointer    =   99  'Custom
      ScaleHeight     =   10320
      ScaleWidth      =   13980
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   540
      Width           =   14040
      Begin VB.CheckBox chkSegnaPc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "&p:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         MousePointer    =   1  'Arrow
         TabIndex        =   25
         ToolTipText     =   "Afficher les points de segmentation de la Spline"
         Top             =   60
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   1560
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label zLabel11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4620
         TabIndex        =   29
         Top             =   60
         Width           =   195
      End
      Begin VB.Label lblY 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4860
         TabIndex        =   28
         Top             =   60
         Width           =   735
      End
      Begin VB.Label zLabel10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "x:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3540
         TabIndex        =   27
         Top             =   60
         Width           =   195
      End
      Begin VB.Label lblX 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3780
         TabIndex        =   26
         Top             =   60
         Width           =   735
      End
      Begin VB.Shape shpPi 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   255
         Index           =   0
         Left            =   2640
         Shape           =   1  'Square
         Top             =   2040
         Width           =   255
      End
      Begin VB.Line linPi 
         BorderColor     =   &H000080FF&
         Index           =   0
         X1              =   3720
         X2              =   1920
         Y1              =   3240
         Y2              =   1200
      End
   End
   Begin MSComctlLib.TabStrip tabTypeC 
      Height          =   10935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   19288
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Be&zier"
            Key             =   "Bezier"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&B-Spline"
            Key             =   "B-Spline"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&C-Spline"
            Key             =   "C-Spline"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&T-Spline"
            Key             =   "T-Spline"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame zFrame2 
      Caption         =   "Echelle du graphique:"
      Height          =   1935
      Left            =   14520
      TabIndex        =   22
      Top             =   9240
      Width           =   3255
      Begin VB.TextBox txtZmin 
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         ToolTipText     =   "OK ou Enter pour confirmer"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtZmax 
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         ToolTipText     =   "OK ou Enter pour confirmer"
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdSG_OK 
         Caption         =   "OK"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtYmax 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         ToolTipText     =   "OK ou Enter pour confirmer"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtYmin 
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         ToolTipText     =   "OK ou Enter pour confirmer"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtXmax 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "OK ou Enter pour confirmer"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtXmin 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "OK ou Enter pour confirmer"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label zLabel09 
         BackStyle       =   0  'Transparent
         Caption         =   "Zma&x"
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   840
         Width           =   615
      End
      Begin VB.Label zLabel08 
         BackStyle       =   0  'Transparent
         Caption         =   "Zmi&n"
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.Label zLabel06 
         BackStyle       =   0  'Transparent
         Caption         =   "Ym&ax"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.Label zLabel05 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ymin"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label zLabel04 
         BackStyle       =   0  'Transparent
         Caption         =   "X&max"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.Label zLabel03 
         BackStyle       =   0  'Transparent
         Caption         =   "&Xmin"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame zFrame1 
      Caption         =   "Paramètre de la courbe:"
      Height          =   1575
      Left            =   14520
      TabIndex        =   21
      Top             =   7680
      Width           =   2295
      Begin VB.CommandButton cmdKZ_OK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   720
         Picture         =   "Splines.frx":0660
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1170
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txtKZ 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "OK ou Enter pour confirmer"
         Top             =   1170
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtNPC 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   "OK ou Enter pour confirmer"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdNPC_OK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1920
         Picture         =   "Splines.frx":0B92
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   270
      End
      Begin VB.TextBox txtNPI 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "OK ou Enter pour confirmer"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdNPI_OK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   720
         Picture         =   "Splines.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   270
      End
      Begin VB.Label lblKZ 
         BackStyle       =   0  'Transparent
         Caption         =   "Pa&ramètre         NK ou VZ"
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label zLabel01 
         BackStyle       =   0  'Transparent
         Caption         =   "N. de points &sur la courbe:"
         Height          =   465
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label zLabel02 
         BackStyle       =   0  'Transparent
         Caption         =   "N. de point a &interpoler:"
         Height          =   465
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrillePoint 
      Height          =   2265
      Left            =   14520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   3995
      _Version        =   393216
      Cols            =   4
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      Enabled         =   -1  'True
      ScrollBars      =   2
      MousePointer    =   99
      MouseIcon       =   "Splines.frx":15F6
   End
   Begin VB.Label zLabel07 
      BackStyle       =   0  'Transparent
      Caption         =   "Poles de la Spline:"
      Height          =   255
      Left            =   14640
      TabIndex        =   23
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Menu mnuCurve 
      Caption         =   "Cour&be"
      Begin VB.Menu mnuChargerPoint 
         Caption         =   "&Charger un fichier de point"
      End
      Begin VB.Menu mnuSauverPoint 
         Caption         =   "&Sauvegarder un fichier de point"
      End
      Begin VB.Menu zSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortie 
         Caption         =   "&Sortie"
      End
   End
End
Attribute VB_Name = "frmSplines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
' Description.....: Programme de visualisation d'interpolation Spline
' Systeme.........: Visual Basic 6.0 sous Windows NT.
' Auteur original.: F. Languasco ®
' Note ...........: Modifié par Cuq pour une visualisation en 3D
'===========================================================================
'
Option Explicit
'
Dim NPI&            ' N. de Points dans la courbe.
Dim Pi() As P_Type  ' Coordonnees des Points de l'interpolation.
Dim NPC&            ' N. de point approximant la courbe.
Dim Pc() As P_Type  ' Coordonnees des points pour l'approximation.
Dim NK&             ' Degree pour la B-Spline.
Dim VZ&             ' Tension de la courbe T-Spline.
'
Dim TypeC$          ' Type d'interpolation activée
'
Dim ShOx!, ShOy!    ' Offset pour le centre de l'indicateur du point ( dépend de l'échelle)
                    
Dim PSel&           ' Indice du point selectionné
'
Dim Xmin!, Xmax!    ' Coordonnees minimum et maximum
Dim Ymin!, Ymax!    ' du quadrillage.
Dim Zmin!, Zmax!    '

Dim Vue             ' Vue actuelle de visualisation
                    ' 0 Vue XY
                    ' 1 Vue XZ
                    ' 2 Vue YZ
Dim ResteValue     ' Sauvegarde de la valeur de l'axe 3D non traité selon la vue
                   ' après Modif dans la grille
'
Dim DirNome$        ' Repertoire des fichiers Splines.
Const PExt$ = "dat" ' Extension des fichiers de point
'
'
Const BZ$ = ""
Const CS$ = ""
Const BS$ = "&Degree" & vbNewLine & "2 <= NK <= NPI"
Const TS$ = "Te&nsion" & vbNewLine & "1 <= VZ <= 100"
'
Dim RS1&                ' Position pour l'edition
Dim CS1&                ' de la coordonnees des points
Dim RS1_O&              ' dans la table
Dim CS1_O&              '
'
Dim GrillePoint_Left&      '
Dim GrillePoint_Top&       '
'
Dim NoPaint As Boolean  ' Evite de redessiner la courbe si il n'y a pas de modification
'
Const PCHL& = &HC0FFFF  ' Couleur  de fond pour la valeur actuelle de la position  du curseur.
'
'--- GetLocale: ----------------------------------------------------------------
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
(ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String _
, ByVal cchData As Long) As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
'
Private Const LOCALE_SDECIMAL& = &HE
Private Const LOCALE_STHOUSAND& = &HF
Private Const LOCALE_SDATE& = &H1D
Private Const LOCALE_STIME& = &H1E
Private Sub GetLocale(Optional ByRef DS$, Optional ByRef MS$ _
    , Optional ByRef GS$, Optional ByRef TS$)
'
'   Trouver les séparateurs suivant la configuration des paramètres régionnaux de Windows:
'    DS$:   séparateur de la decimale.
'    MS$:   séparateur des milliers.
'    GS$:   séparateur pour les dates.
'    TS$:   séparateur pour les heures.
'
    DS$ = "  "
    MS$ = "  "
    GS$ = "  "
    TS$ = "  "
'
    GetLocaleInfo GetThreadLocale(), LOCALE_SDECIMAL, DS$, Len(DS$)
    GetLocaleInfo GetThreadLocale(), LOCALE_STHOUSAND, MS$, Len(MS$)
    GetLocaleInfo GetThreadLocale(), LOCALE_SDATE, GS$, Len(GS$)
    GetLocaleInfo GetThreadLocale(), LOCALE_STIME, TS$, Len(TS$)
'
    DS$ = Left$(DS$, 1)
    MS$ = Left$(MS$, 1)
    GS$ = Left$(GS$, 1)
    TS$ = Left$(TS$, 1)
'
'
'
End Sub
Private Function KAscNumInteri(ByVal KA%, Optional ByVal NEG As Boolean = False) As Integer
'
' Permet de filter les valeurs et n'accepte que des valeurs numérique dans le cas de texteBox
' Retourne uniquement des valeurs Entière
'
    Dim KeyMinus%
    Dim TextB As TextBox    ' Pour type TextBoxes.
'
    Const myKeyMinus% = 45  ' ' Valeur retournée après appui sur la touche '-'
'
    Set TextB = Screen.ActiveControl
'
    ' Filtre pour le signe "-":
    If (Left$(TextB.Text, 1) <> "-" Or TextB.SelText = TextB.Text) _
    And NEG And TextB.SelStart = 0 Then KeyMinus = myKeyMinus
'
    Select Case KA
        Case vbKey0 To vbKey9, KeyMinus, vbKeyBack
        KAscNumInteri = KA
'
        Case Else
        KAscNumInteri = 0
    End Select
'
'
'
End Function

Private Function KAscNumReali(ByVal KA As Integer, Optional ByVal NEG As Boolean = False) As Integer
'
' Permet de filter les valeurs et n'accepte que des valeurs numérique dans le cas de texteBox
' Retourne uniquement des valeurs Réelle en notation scientifique   1 E+99
'
    Dim KeyDecimal%, KeyMinus%, KeyE%
    Dim TextB As TextBox    ' Pour type TextBoxes.

    Dim SD$, SM$, myKeyDecimal%
'
    Const myKeyMinus% = 45  ' ' Valeur retournée après appui sur la touche '-'
    
    GetLocale SD$, SM$          ' Récupère les paramètres régionnaux pour les séparateurs
    myKeyDecimal% = Asc(SD$)    ' Séparateur des milliers
'
    Set TextB = Screen.ActiveControl
'
    ' Filtre pour la décimale:
    If (InStr(TextB.Text, SD$) = 0 _
    And Not (TextB.SelStart = 0 And Left$(TextB.Text, 1) = "-")) _
    Or TextB.SelText = TextB.Text Then KeyDecimal = myKeyDecimal
'
    ' Filtre pour le signe "-":
    If (Left$(TextB.Text, 1) <> "-" Or TextB.SelText = TextB.Text) _
    And NEG And TextB.SelStart = 0 Then KeyMinus = myKeyMinus
'
    ' Filtre pour la notation scientifique (1 E+99 ):
    If TextB.SelStart > 0 Then
        KA = Asc(UCase$(Chr$(KA)))
        If (InStr(TextB.Text, "E") = 0 _
        And Not (TextB.SelStart = 0 Or TextB.SelText = TextB.Text)) _
        And Mid$(TextB.Text, TextB.SelStart, 1) <> "-" Then KeyE = vbKeyE
'
        If Mid$(TextB.Text, TextB.SelStart, 1) = "E" Then KeyMinus = myKeyMinus
'
        If (InStr(TextB.Text, "E") > 0) _
        And (TextB.SelStart - InStr(TextB.Text, "E") >= 0) Then KeyDecimal = 0
    End If
'
    Select Case KA
        Case vbKey0 To vbKey9, KeyDecimal, KeyMinus, vbKeyBack, KeyE
        KAscNumReali = KA
'
        Case Else
        KAscNumReali = 0
    End Select
'
'
'
End Function
Private Sub chkSegnaPc_Click()
'
'
    picQuadrillage.Cls
    DessinCourbe picQuadrillage
'
    chkSegnaPc.ToolTipText = Switch( _
    chkSegnaPc = vbUnchecked, "Affiche les points d'approximation de la spline", _
    chkSegnaPc = vbChecked, "Masque les points d'approximation de la spline")
'
'
End Sub

Private Sub cmdKZ_OK_Click()
'
' Valide les valeurs selon le cas de la B-Spline ( degrée compris entre 2 et N )
' Valide les valeurs selon le cas de la T-Spline ( Tension compris entre 1 et 100 )

    Select Case TypeC$
        Case "B-Spline"
        If Val(txtKZ) >= 2 And Val(txtKZ) <= NPI Then
            NK = Val(txtKZ)
            picQuadrillage.Refresh
        Else
            txtKZ = NK
        End If
'
        Case "T-Spline"
        If Val(txtKZ) >= 1 And Val(txtKZ) <= 100 Then
            VZ = Val(txtKZ)
            picQuadrillage.Refresh
        Else
            txtKZ = VZ
        End If
    End Select
'
'
'
End Sub

Private Sub cmdNPI_OK_Click()
'
'
    If Val(txtNPI) >= 3 Then
        NPI = Val(txtNPI)
        If NK > NPI Then
            NK = NPI
            If TypeC$ = "B-Spline" Then txtKZ = NK
        End If
        ReDim Preserve Pi(0 To NPI - 1) ' As P_Type
        AjouterPoint
        PositionnerPoint
    Else
        txtNPI = NPI
    End If
'
'
'
End Sub

Private Sub cmdSG_OK_Click()
'
'
    Dim X1!, X2!, Y1!, Y2!, Z1!, Z2!
'
    If VerificaSG(X1!, Y1!, X2!, Y2!, Z1!, Z2!) Then
        Xmin = X1
        Xmax = X2
        Ymin = Y1
        Ymax = Y2
        Zmin = Z1
        Zmax = Z2
'
        Select Case Vue
        Case 0
            Quadrillage picQuadrillage, Xmin, Xmax, Ymin, Ymax, , , 2, ShOx, ShOy, , "x", "y"
        Case 1
            Quadrillage picQuadrillage, Xmin, Xmax, Zmin, Zmax, , , 2, ShOx, ShOy, , "x", "z"
        Case 2
            Quadrillage picQuadrillage, Ymin, Ymax, Zmin, Zmax, , , 2, ShOx, ShOy, , "y", "z"
        End Select
        
        PositionnerPoint
    End If
'
'
'
End Sub

Private Sub cmdNPC_OK_Click()
'
'
    If Val(txtNPC) >= 2 Then
        NPC = Val(txtNPC)
        picQuadrillage.Refresh
    Else
        txtNPC = NPC
    End If
'
'
'
End Sub

Private Sub Form_Load()
'
'
    Dim CMDFilter$
' Initialisation
    NPI = 6     ' N. de Points d'interpolation .
    txtNPI = NPI
    NPC = 100   ' N. de Points d'approximation.
    txtNPC = NPC
'
    NK = 3      ' Degree pour la  B-Spline.
    VZ = 30     ' Tension pour la courbe T-Spline
    
    ' Coordonnees des points de départ:
    ReDim Pi(0 To NPI - 1) ' As P_Type
    Pi(0).X = -10: Pi(0).Y = -3.5: Pi(0).Z = -1
    Pi(1).X = 3: Pi(1).Y = 3: Pi(1).Z = 0
    Pi(2).X = 3: Pi(2).Y = -0.01: Pi(2).Z = 1
    Pi(3).X = 5: Pi(3).Y = 3: Pi(3).Z = 3
    Pi(4).X = 7: Pi(4).Y = 4.2: Pi(4).Z = 0
    Pi(5).X = 10: Pi(5).Y = 4: Pi(5).Z = -4
'
    ' Dimensione le Quadrillage:
    Xmin = -10
    txtXmin = Xmin
    Ymin = -5
    txtYmin = Ymin
    Xmax = 10
    txtXmax = Xmax
    Ymax = 5
    txtYmax = Ymax
    Zmax = 3
    txtZmax = Zmax
    Zmin = -4
    txtZmin = Zmin
    
'Dim DataFileName As String, Xdata As Double, Ydata As Double, numPoints As Long
'Dim MinX As Double, MaxX As Double, filein%, iSkip As Long
'Dim NumOutput As Integer, SkipNum As Integer
'Dim MaxY As Double, MinY As Double
'
'MinX = 999999
'MinY = MinX
'MaxX = -MinX
'MaxY = MaxX
'
'NumOutput = 300
'SkipNum = 2 'only read every SkipNum points in order not to overload the Belzier binomial coeficient
'
'DataFileName = "c:\jk\druk-vangeld-data\drukmix-all-sorted.txt"
'If Dir(DataFileName) <> "" Then
'    filein% = FreeFile
'    Open DataFileName For Input As #filein%
'    numPoints = 0
'    iSkip = 0
'    Do Until EOF(filein%)
'        Input #filein%, Xdata, Ydata
'        iSkip = iSkip + 1
'        If iSkip Mod 2 = 0 Then
'            ReDim Preserve Pi(0 To numPoints)
'            Pi(numPoints).X = Xdata
'            Pi(numPoints).Y = Ydata
'            Pi(numPoints).Z = 0#
'            If Xdata > MaxX Then MaxX = Xdata
'            If Xdata < MinX Then MinX = Xdata
'            If Ydata > MaxY Then MaxY = Ydata
'            If Ydata < MinY Then MinY = Ydata
'            numPoints = numPoints + 1
'            End If
'    Loop
'    Close #filein%
'    NPI = numPoints
'    NPC = NPI 'NumOutput
'
'    txtNPI = NPI
'    txtNPC = NPC
''
'    NK = 3      ' Degree pour la  B-Spline.
'    VZ = 30     ' Tension pour la courbe T-Spline
'
'    Xmin = MinX
'    Xmax = MaxX
'    Ymin = MinY
'    Ymax = MaxY
'    Zmin = 0
'    Zmax = 0
'
'    txtXmin = Xmin
'    txtYmin = Ymin
'    txtXmax = Xmax
'    txtYmax = Ymax
'    txtZmax = Zmax
'    txtZmin = Zmin
'
' Else
'    Call MsgBox("File not found!", vbCritical, "File missing")
'    Exit Sub
'    End If
    
    Quadrillage picQuadrillage, Xmin, Xmax, Ymin, Ymax, , , 2, ShOx, ShOy, , "x", "y"
'
    ' Type d'interpolation initiale:
'    tabTypeC.Tabs("Bezier").Selected = True
    tabTypeC.Tabs("B-Spline").Selected = True
'
    ' Dimmensione l'indicateur de pole:
    shpPi(0).Width = picQuadrillage.ScaleX(6, vbPixels, vbUser)
    shpPi(0).Height = Abs(picQuadrillage.ScaleY(6, vbPixels, vbUser))
'
    'Dessine les poles:
    AjouterPoint
    PositionnerPoint
'
    ' Dimensione la grille de point
    With GrillePoint
        .Cols = 4
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignRightCenter
        .FixedAlignment(0) = flexAlignRightCenter
        .FixedAlignment(1) = flexAlignRightCenter
        .FixedAlignment(2) = flexAlignRightCenter
        .FixedAlignment(3) = flexAlignRightCenter
'
        .ColWidth(0) = 420
        .ColWidth(1) = 810
        .ColWidth(2) = 810
        .ColWidth(3) = 810
'
        .Row = 0
        .Col = 0
        .Text = "I"
        .Col = 1
        .Text = "Pi(I).x"
        .Col = 2
        .Text = "Pi(I).y"
        .Col = 3
        .Text = "Pi(I).z"
'
        GrillePoint_Left = .Left + 45
        GrillePoint_Top = .Top + 45
    End With
'
    ' Initialise les paramètres Pour la sauvegarde et le chargement des points
        DirNome$ = App.Path
    CMDFilter$ = "Dati per Splines (*." & PExt$ & ")|*." & PExt$
    CMDFilter$ = CMDFilter$ & "|Tutti i Files (*.*)|*.*"
    CMDialog1.Filter = CMDFilter$
    ' Cahe les fichiers en lecture seul , controle l'existence Nasconde casella Read Only, controlla esistenza Files
    ' e avverte per sovrascrittura Files:
    CMDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNOverwritePrompt
'
'
'
End Sub
Private Sub MiseAJourPositionCurseur(ByVal sVal1$, ByVal sVal2$)
'
'   Mise a jour de la position du curseur
'
    lblX = sVal1$
    lblY = sVal2$
'
    lblX.BackColor = IIf(sVal1$ = "", vbButtonFace, PCHL)
    lblY.BackColor = IIf(sVal2$ = "", vbButtonFace, PCHL)
'
'
'
End Sub

Private Sub PositionnerPoint()
'
'   Posiziona, sul grafico, Shapes e Lines:
'
    Dim I&
'
    NoPaint = True
    
    Select Case Vue
    Case 0  ' Vue XY
            For I = 0 To NPI - 2
                shpPi(I).Left = Pi(I).X - ShOx
                shpPi(I).Top = Pi(I).Y + ShOy
                linPi(I).X1 = Pi(I).X
                linPi(I).Y1 = Pi(I).Y
                linPi(I).X2 = Pi(I + 1).X
                linPi(I).Y2 = Pi(I + 1).Y
            Next I
            shpPi(NPI - 1).Left = Pi(NPI - 1).X - ShOx
            shpPi(NPI - 1).Top = Pi(NPI - 1).Y + ShOy
            
    Case 1 ' Vue XZ
            For I = 0 To NPI - 2
                shpPi(I).Left = Pi(I).X - ShOx
                shpPi(I).Top = Pi(I).Z + ShOy
                linPi(I).X1 = Pi(I).X
                linPi(I).Y1 = Pi(I).Z
                linPi(I).X2 = Pi(I + 1).X
                linPi(I).Y2 = Pi(I + 1).Z
            Next I
            shpPi(NPI - 1).Left = Pi(NPI - 1).X - ShOx
            shpPi(NPI - 1).Top = Pi(NPI - 1).Z + ShOy
    Case 2 ' Vue YZ
            For I = 0 To NPI - 2
                shpPi(I).Left = Pi(I).Y - ShOx
                shpPi(I).Top = Pi(I).Z + ShOy
                linPi(I).X1 = Pi(I).Y
                linPi(I).Y1 = Pi(I).Z
                linPi(I).X2 = Pi(I + 1).Y
                linPi(I).Y2 = Pi(I + 1).Z
            Next I
            shpPi(NPI - 1).Left = Pi(NPI - 1).Y - ShOx
            shpPi(NPI - 1).Top = Pi(NPI - 1).Z + ShOy
            
    End Select
    
'
    NoPaint = False
    picQuadrillage.Refresh
'
'
'
End Sub
Private Function TrouverPoint(ByVal X!, ByVal Y!) As Long
'
'  Retourne l'indice du point
    Dim I&, Xc!, Yc!, DsQ!, DisQ!, ShOx2!, ShOy2!
'
    ShOx2 = ShOx ^ 2
    ShOy2 = ShOy ^ 2
    
    DisQ = 1E+38
'
    For I = 0 To NPI - 1
        Xc = shpPi(I).Left + ShOx
        Yc = shpPi(I).Top - ShOy
        
        DsQ = (X - Xc) * (X - Xc) / ShOx2 + (Y - Yc) * (Y - Yc) / ShOy2
        If DisQ > DsQ Then
            DisQ = DsQ
            TrouverPoint = I
        End If
    Next I
'
'
'
End Function


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    MiseAJourPositionCurseur "", ""
'
'
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'
    End
'
'
'
End Sub

Private Sub GrillePoint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    RS1 = GrillePoint.Row
    CS1 = GrillePoint.Col
    txtEditPoint.Visible = False
'
'
'
End Sub


Private Sub GrillePoint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    GrillePoint.Row = RS1
    GrillePoint.Col = CS1
    txtEditPoint.Text = GrillePoint.Text
'
    txtEditPoint.Left = GrillePoint.ColPos(CS1) + GrillePoint_Left
    txtEditPoint.Top = GrillePoint.RowPos(RS1) + GrillePoint_Top
    txtEditPoint.Width = GrillePoint.ColWidth(CS1) - 15
    txtEditPoint.Height = GrillePoint.RowHeight(RS1) - 15
    txtEditPoint.Visible = True
    txtEditPoint.SetFocus
'
    RS1_O = RS1
    CS1_O = CS1
'
'
'
End Sub






Private Sub mnuSortie_Click()
'
'
    End
'
'
'
End Sub



Private Sub mnuChargerPoint_Click()
'
'
    Dim FN_Temp$, M$
'
    On Error GoTo mnuChargerPoint_ERR
'
    CMDialog1.DialogTitle = " Charger les points d'interpolation"
    CMDialog1.InitDir = DirNome$
    CMDialog1.ShowOpen
    FN_Temp$ = CMDialog1.FileName
'
    If BreakDown(FN_Temp$, DirNome$) Then
        ChargerPoint FN_Temp$
    End If
'
    Exit Sub
'
'
mnuChargerPoint_ERR:
    If Err <> cdlCancel Then
        M$ = "Erreur " & Str$(Err.Number) & vbNewLine
        M$ = M$ & Err.Description
        MsgBox M$, vbCritical, "Erreur dans mnuChargerPoint " & Err.Source
    End If
'
'
'
End Sub

Private Sub mnuSauverPoint_Click()
'
'
    Dim FN_Temp$, M$
'
    On Error GoTo mnuSauverPoint_ERR
'
    CMDialog1.DialogTitle = " Sauver les points de la Spline"
    CMDialog1.FileName = "*." & PExt$
    CMDialog1.InitDir = DirNome$
    CMDialog1.ShowSave
    FN_Temp$ = CMDialog1.FileName
'
    BreakDown FN_Temp$, DirNome$
    SauverPoint FN_Temp$
'
    Exit Sub
'
'
mnuSauverPoint_ERR:
    If Err <> cdlCancel Then
        M$ = "Erreur " & Str$(Err.Number) & vbNewLine
        M$ = M$ & Err.Description
        MsgBox M$, vbCritical, "Erreur dans mnuSauverPoint " & Err.Source
    End If
'
'
'
End Sub
Private Function BreakDown(ByVal Full$, Optional ByRef PName$ _
    , Optional ByRef FName$, Optional ByRef Ext$) As Boolean
'
'   Décompose un nom de fichier en différente sous partie:
'   Full$  = Nom Complet du fichier.
'   PName$ = Chemin du fichier.
'   FName$ = Nom du fichier avec son extension.
'   Ext$   = .extension du fichier.
'
'   Si le fichier n'existe pas retourne une valeur False.
'
    Dim Sloc&, Dot&
'
    BreakDown = Len(Dir$(Full$))
'
    If InStr(Full$, "\") Then
        FName$ = Full$
        PName$ = ""
        Sloc = InStr(FName$, "\")
        Do While Sloc <> 0
            PName$ = PName$ & Left$(FName$, Sloc)
            FName$ = Mid$(FName$, Sloc + 1)
            Sloc = InStr(FName$, "\")
        Loop
    Else
        PName$ = ""
        FName$ = Full$
    End If
'
    Dot = InStr(Full$, ".")
    If Dot <> 0 Then
        Ext$ = Mid$(Full$, Dot)
    Else
        Ext$ = ""
    End If
'
'
'
End Function

Private Function Arrond(ByVal X As Double) As Double
'
'   Arrondi une valeur réelle a l'entier supérieur:
' Permet dàvoir par exemple pour l'echelle de représentation deux valeurs non équivalente
'
    If X = Int(X) Then
        Arrond = X
    Else
        Arrond = Int(X) + 1#
    End If
'
'
'
End Function

Private Sub OptionVue_Click(Index As Integer)

' Mise a jour de la vue actuelle
Vue = Index


' Selection de la vue

Select Case Index

Case 0 ' Vue XY

    zLabel10 = "X:"
    zLabel11 = "Y:"

    Quadrillage picQuadrillage, Xmin, Xmax, Ymin, Ymax, , , 2, ShOx, ShOy, , "x", "y"


Case 1 ' Vue XZ

    zLabel10 = "X:"
    zLabel11 = "Z:"
    
    Quadrillage picQuadrillage, Xmin, Xmax, Zmin, Zmax, , , 2, ShOx, ShOy, , "x", "z"

    
Case 2 ' Vue YZ

    zLabel10 = "Y:"
    zLabel11 = "Z:"
    
    Quadrillage picQuadrillage, Ymin, Ymax, Zmin, Zmax, , , 2, ShOx, ShOy, , "y", "z"


End Select

       

PositionnerPoint

End Sub

Private Sub picQuadrillage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    If Button = vbLeftButton Then
        PSel = TrouverPoint(X, Y)
        picQuadrillage_MouseMove Button, Shift, X, Y
    End If
'
'
'
End Sub


Private Sub picQuadrillage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    If Button = vbLeftButton Then
        NoPaint = True
        shpPi(PSel).Move X - ShOx, Y + ShOy

'

        If PSel = 0 Then
            linPi(0).X1 = X
            linPi(0).Y1 = Y
        ElseIf PSel = NPI - 1 Then
            linPi(PSel - 1).X2 = X
            linPi(PSel - 1).Y2 = Y
        Else
            linPi(PSel).X1 = X
            linPi(PSel).Y1 = Y
            linPi(PSel - 1).X2 = X
            linPi(PSel - 1).Y2 = Y
        End If
        
        
        NoPaint = False
'
        
        Select Case Vue
        Case 0 ' Vue XY
            Pi(PSel).X = X
            Pi(PSel).Y = Y
            Pi(PSel).Z = ResteValue
        Case 1 'VUE XZ
            Pi(PSel).X = X
            Pi(PSel).Y = ResteValue
            Pi(PSel).Z = Y
        Case 2 ' VUE YZ
            Pi(PSel).X = ResteValue
            Pi(PSel).Y = X
            Pi(PSel).Z = Y
        End Select
        
        
        ' Mise a jour de la grille de point
        With GrillePoint
            .Row = PSel + 1
            .Col = 1
            .Text = Format$(Pi(PSel).X, "#0.000")
            .Col = 2
            .Text = Format$(Pi(PSel).Y, "#0.000")
            .Col = 3
            .Text = Format$(Pi(PSel).Z, "#0.000")
        End With
'
    ElseIf Button = 0 Then
        picQuadrillage.ToolTipText = ""
        picQuadrillage.ToolTipText = "Bouger le point " & TrouverPoint(X, Y)
    End If
'
    MiseAJourPositionCurseur Format$(X, "#0.###"), Format$(Y, "#0.###")
'
'
'
'
'
'
End Sub


Private Sub picQuadrillage_Paint()
'
'
    picQuadrillage.Cls
'
    If Not NoPaint Then DessinCourbe picQuadrillage
'
'
'
End Sub


Private Sub tabTypeC_Click()
'
'
    TypeC$ = tabTypeC.SelectedItem.Key
'
    Select Case TypeC$
        Case "Bezier"
        txtKZ.Visible = False
        cmdKZ_OK.Visible = False
        lblKZ = BZ$
        
'
        Case "B-Spline"
        txtKZ = NK
        txtKZ.Visible = True
        cmdKZ_OK.Visible = True
        lblKZ = BS$
'
        Case "C-Spline"
        txtKZ.Visible = False
        cmdKZ_OK.Visible = False
        lblKZ = CS$
'
        Case "T-Spline"
        txtKZ = VZ
        txtKZ.Visible = True
        cmdKZ_OK.Visible = True
        lblKZ = TS$
    End Select
'
    picQuadrillage.Refresh
'
'
'
End Sub

Private Sub tabTypeC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    MiseAJourPositionCurseur "", ""
'
'
'
End Sub

Private Sub txtEditPoint_KeyPress(KeyAscii As Integer)
'
'
    Select Case KeyAscii
        Case vbKeyEscape
        txtEditPoint.Text = ""
        txtEditPoint.Visible = False
        KeyAscii = 0
'
        Case vbKeyReturn
        txtEditPoint.Visible = False
        KeyAscii = 0
'
        Case Else
        KeyAscii = KAscNumReali(KeyAscii, True)
    End Select
'
'
'
End Sub
Private Sub txtEditPoint_LostFocus()
'
'
    Dim X!, Y!, Z!
'
    If txtEditPoint.Text <> "" Then
        GrillePoint.Row = RS1_O
        PSel = RS1_O - 1
'
        Select Case CS1_O
            Case 1 ' Edition X
            
            X = Val(txtEditPoint.Text)
            GrillePoint.Col = 2
            Y = Val(GrillePoint.Text)
            GrillePoint.Col = 3
            Z = Val(GrillePoint.Text)
            
'
            Case 2 '' Edition Y
            GrillePoint.Col = 1
            X = Val(GrillePoint.Text)
            Y = Val(txtEditPoint.Text)
            GrillePoint.Col = 3
            Z = Val(GrillePoint.Text)
            
            Case 3 ' Edition Z
            
            GrillePoint.Col = 1
            X = Val(GrillePoint.Text)
            GrillePoint.Col = 2
            Y = Val(GrillePoint.Text)
            Z = Val(txtEditPoint.Text)
            
        End Select
'


Select Case Vue
Case 0
        ResteValue = Z
        picQuadrillage_MouseMove vbLeftButton, 0, X, Y
Case 1
        ResteValue = Y
        picQuadrillage_MouseMove vbLeftButton, 0, X, Z
Case 2
        ResteValue = X
        picQuadrillage_MouseMove vbLeftButton, 0, Y, Z
End Select

    End If
'
    txtEditPoint.Visible = False
    txtEditPoint.Text = ""
'
'
'
End Sub


Private Sub txtKZ_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then cmdKZ_OK_Click
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub


Private Sub txtNPI_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then cmdNPI_OK_Click
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub


Private Sub txtNPC_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then cmdNPC_OK_Click
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub



Private Function Quadrillage(ByVal Page As PictureBox _
    , ByVal X0!, ByVal Xn!, ByVal Y0!, ByVal Yn! _
    , Optional ByVal FormatVX$ = "#0.0##" _
    , Optional ByVal FormatVY$ = "#0.0##" _
    , Optional ByVal Npx& = 1, Optional PxN_X!, Optional PxN_Y! _
    , Optional ByVal Titre$ = "" _
    , Optional ByVal UniteX$ = "" _
    , Optional ByVal UniteY$ = "" _
    , Optional ByVal AutoRed As Boolean = False) As Boolean
'
'   Routine permettant d'initialise un controle de type Picture box
'   pour la représentation d'une fonction du type y = f(x).
'    Page:    PictureBox de destination.
'    X0:        Valeur minimale de l'axe des absices.
'    Xn:        Valeur maximale de l'axe des absices.
'    Y0:        Valeur minimale de l'axe des ordonnees.
'    Yn:        Valeur maximale de l'axe des ordonnees.
'    FormatVX$: Chaine de format des valeur sur l'axe X.
'    FormatVY$: Chaine de format des valeur sur l'axe Y.
'    Npx:       N° de Pixels de cui si vuole conoscere
'    PxN_X:      la largeur in [vbUser] e
'    PxN_Y:      la hauteur in [vbUser].
'    Titre$:   Titre du graphique.
'    UniteX$:   Unite' (ou titre) de l' axe X.
'    UniteY$:   Unite' (ou titre) de l' axe Y.
'    AutoRed:   Sauvegarde de Page.AutoRedraw pour le dessin du quadrillage.
'
    Dim I&, Xi!, D_X!, rrx!, Yi!, D_Y!, rry!, Tx$
    Dim QxMin!, QxMax!, QyMin!, QyMax!, QzMin!, QzMax!, b0!, bn!, TxW!
    Dim TitL!, TitT!, TitW!, TitH!, Po4_X!, Po4_Y!
    Const Log10! = 2.30258509299405 ' Log(10#)
    Const DYMin! = 0.0001           ' Valeur min. du facteur d'echelle Y.
'
    On Error GoTo Quadrillage_ERR
    ' Controle de l'echelle
    If X0 >= Xn Then Err.Raise 1001, "Quadrillage", "Erreur de l'echelle sur X."
    If Y0 > Yn Then Err.Raise 1001, "Quadrillage", "Erreur de l'echelle sur Y."
'
    ' Impose les valeurs pour les charactères de Font des valeurs
    ' des différents axes:
    Page.FontName = "MS Sans Serif"
    Page.FontSize = 8
    Page.FontBold = False
'
    ' Calcul la interval des valeurs écrites
    ' sur l'axe X: la sequence est  1, 2, 2.5 e 5:
    D_X = Xn - X0
    rrx = 10! ^ Arrond(Log(D_X / 20!) / Log10)
    Do While D_X / rrx < 5!
        rrx = rrx / 2!
    Loop
    If D_X / rrx > 10! Then rrx = rrx * 2!
    X0 = rrx * Int(Round(X0 / rrx, 3))
    Xn = rrx * Arrond(Round(Xn / rrx, 3))
    D_X = Xn - X0
'
    ' Impose un facteur minimum
    ' pour l' axe Y:
    If Yn - Y0 < DYMin Then
        Y0 = Y0 - DYMin / 2!
        Yn = Yn + DYMin / 2!
    End If
'
    ' Calcul l'interval des valeurs écrites
    ' sur l'axe Y: la sequence est  1, 2, 2.5 e 5:
    D_Y = Yn - Y0
    rry = 10! ^ Arrond(Log(D_Y / 20!) / Log10)
    Do While D_Y / rry < 5!
        rry = rry / 2!
    Loop
    If D_Y / rry > 10! Then rry = rry * 2!
    Y0 = rry * Int(Round(Y0 / rry, 3))
    Yn = rry * Arrond(Round(Yn / rry, 3))
    D_Y = Yn - Y0
'
    ' la bordure a droite dépend de la
    ' presence, ou non, d'une étiquette:
    If UniteX$ = "" Then
        bn = D_X / 20!
    Else
        bn = D_X / 10!
    End If
'
    ' la bordure a gauche doit etre suffisante
    ' pour contenir la valeur de Y la plus large:
    TxW = Page.TextWidth(Format$(Y0, FormatVY$) & " ")
    If TxW < Page.TextWidth(Format$(Yn, FormatVY$) & " ") Then
        TxW = Page.TextWidth(Format$(Yn, FormatVY$) & " ")
    End If
    b0 = TxW * (D_X + bn) / (Page.ScaleWidth - TxW)
    If b0 < D_X / 10! Then b0 = D_X / 10!
'
    ' Impose la bordure horizontale
    ' et verticale:
    QxMin = X0 - b0
    QxMax = Xn + bn
    QyMin = Y0 - D_Y / 10!
    QyMax = Yn + D_Y / 7!
'
    ' Annule 'limage et impose l'echelle:
    Page.Picture = LoadPicture("")
    Page.Scale (QxMin, QyMax)-(QxMax, QyMin)
    ' le dessin dans l'image n'est pas effacé:
    Page.AutoRedraw = True
    ' Calcul la largeur et la hauteur de Npx pixels:
    PxN_X = Abs(Page.ScaleX(Npx, vbPixels, vbUser))
    PxN_Y = Abs(Page.ScaleY(Npx, vbPixels, vbUser))
    ' Calcul la largeur et la hauteur de 4 points:
    Po4_X = Page.ScaleX(4, vbPoints, vbUser)
    Po4_Y = Page.ScaleY(4, vbPoints, vbUser)
'
    Page.DrawMode = vbCopyPen
    Page.DrawWidth = 1
    Page.DrawStyle = vbDash
    Page.ForeColor = vbGreen
    ' Tracer la grille verticale et ecriture de
    ' la valeur de l' axe X:
    For Xi = X0 To Xn + 0.1 * rrx Step rrx
        Page.Line (Xi, Y0)-(Xi, Yn), vbGreen
        Tx$ = Format$(Xi, FormatVX$)
        ' Verification du format pour éviter une erreur
         ' lors de l'affichage:
        If Abs(Xi - Val(Tx$)) < rrx / 10 Then
            Page.CurrentX = Xi - Page.TextWidth(Tx$) / 2!
            Page.CurrentY = Y0 - D_Y / 70!
            Page.Print Tx$;
        End If
    Next Xi
    ' Ecriture de l' étiquette de l' axe X:
    If UniteX$ <> "" Then
        ' étiquette toute à droite:
        ' Page.CurrentX = QxMax - Page.TextWidth(UniteX$ & " ")
        ' étiquette centrée selon la plus grande valeur et la bordure a droite:
        Page.CurrentX = (Page.CurrentX + QxMax - Page.TextWidth(UniteX$)) / 2!
        Page.Print UniteX$;
    End If
    ' Tracer l' axe Y:
    If (X0 <= 0!) And (0! <= Xn) Then
        Page.DrawStyle = vbSolid
        Page.Line (0!, Y0)-(0!, QyMax - D_Y / 30!), vbGreen
        Page.Line (0!, QyMax - D_Y / 30!) _
                   -(-Po4_X / 2!, Po4_Y + QyMax - D_Y / 30!), vbGreen
        Page.Line (0!, QyMax - D_Y / 30!) _
                   -(Po4_X / 2!, Po4_Y + QyMax - D_Y / 30!), vbGreen
    End If
'
    Page.DrawStyle = vbDash
    ' Tracer la grille horizontale et ecriture de
    ' la valeur de l' axe Y:
    For Yi = Y0 To Yn + 0.1 * rry Step rry
        Page.Line (X0, Yi)-(Xn, Yi), vbGreen
        Tx$ = Format$(Yi, FormatVY$)
        Page.CurrentX = QxMin
        Page.CurrentY = Yi - Page.TextHeight(Tx$) / 2!
        Page.Print Tx$;
    Next Yi
    ' Ecriture de l' étiquette de l' axe Y:
    If UniteY$ <> "" Then
        Page.CurrentX = QxMin
        Page.CurrentY = QyMax
        Page.Print UniteY$;
    End If
    ' Tracer l' axe X:
    If (Y0 <= 0!) And (0! <= Yn) Then
        Page.DrawStyle = vbSolid
        Page.Line (X0, 0!)-(QxMax - D_X / 30!, 0!), vbGreen
        Page.Line (QxMax - D_X / 30!, 0!) _
                   -(QxMax - D_X / 30! - Po4_X, -Po4_Y / 2!), vbGreen
        Page.Line (QxMax - D_X / 30!, 0!) _
                   -(QxMax - D_X / 30! - Po4_X, Po4_Y / 2!), vbGreen
    End If
'
    ' Ecriture du titre du graphique :
    If Titre$ <> "" Then
        Page.FontSize = 12
        Page.FontBold = True
        Page.ForeColor = vbRed
'
        TitW = Page.TextWidth(Titre$)
        TitH = Page.TextHeight(Titre$)
        ' Verification que le titre tient dans toute la Page:
        If TitW <= Page.ScaleWidth Then
            TitL = (QxMin + QxMax - TitW) / 2!
        ' Si il déborde:
        Else
            TitL = Page.ScaleLeft
            Tx$ = " . . . ."
            Titre$ = Left$(Titre$, Int(Len(Titre$) * _
            (Page.ScaleWidth - Page.TextWidth(Tx$)) / TitW)) & Tx$
        End If
        TitT = QyMax
        ' Annul la zonne dans laquelle le titre est écrit ( fond de la couleur du controle:
        'Page.Line (TitL, TitT)-(TitL + TitW, TitT + TitH), Page.BackColor, BF
        Page.CurrentX = TitL
        Page.CurrentY = TitT
        Page.Print Titre$
    End If
'
    Page.DrawStyle = vbSolid
    Page.AutoRedraw = AutoRed
'
'
Quadrillage_ERR:
    Quadrillage = (Err = 0)
    If Err <> 0 Then
        MsgBox Err.Description, vbCritical, Err.Source
    End If
'
'
'
End Function



Private Sub SauverPoint(ByVal FileNome$)
'
'   Sauvegarde dans un fichier texte les points de la courbe
'
    Dim FF%, I&
'
    FF = FreeFile
    Open FileNome For Output As #FF
'
    For I = 0 To NPC - 1
        Write #FF, Pc(I).X, Pc(I).Y, Pc(I).Z
    Next I
'
    Close #FF
'
'
End Sub

Private Sub txtXmax_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then cmdSG_OK_Click
    KeyAscii = KAscNumReali(KeyAscii, True)
'
'
'
End Sub


Private Sub txtXmin_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then cmdSG_OK_Click
    KeyAscii = KAscNumReali(KeyAscii, True)
'
'
'
End Sub


Private Sub txtYmax_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then cmdSG_OK_Click
    KeyAscii = KAscNumReali(KeyAscii, True)
'
'
'
End Sub


Private Sub txtYmin_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then cmdSG_OK_Click
    KeyAscii = KAscNumReali(KeyAscii, True)
'
'
'
End Sub

Private Sub txtZmax_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then cmdSG_OK_Click
    KeyAscii = KAscNumReali(KeyAscii, True)
'
'
'
End Sub

Private Sub txtZmin_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then cmdSG_OK_Click
    KeyAscii = KAscNumReali(KeyAscii, True)
'
'
'
End Sub


Private Function VerificaSG(X1!, Y1!, X2!, Y2!, Z1!, Z2!) As Boolean
'
'  Verification de l'echelle graphique utilisée
'
    Dim M$
'
    On Error Resume Next
'
    If CSng(txtXmin) >= CSng(txtXmax) Then
        M$ = "Xmin ou Xmax Erreur" & vbNewLine
    End If
    X1 = CSng(txtXmin)
    X2 = CSng(txtXmax)
'
    If CSng(txtYmin) >= CSng(txtYmax) Then
        M$ = M$ & "Ymin ou Ymax Erreur" & vbNewLine
    End If
    Y1 = CSng(txtYmin)
    Y2 = CSng(txtYmax)
'
    If CSng(txtZmin) >= CSng(txtZmax) Then
        M$ = M$ & "Zmin ou Zmax Erreur" & vbNewLine
    End If
    Z1 = CSng(txtZmin)
    Z2 = CSng(txtZmax)
    
'
    If M$ <> "" Then MsgBox M$, vbCritical _
    , " Erreur dans les parametres du graphique"
    VerificaSG = (M$ = "")
'
'
'
End Function

Private Sub AjouterPoint()
'
'   Mise a jour du tableau de point
'
    Dim I&
'
    If NPI > shpPi.Count Then
        For I = shpPi.Count To NPI - 1
            Load shpPi(I)
            shpPi(I).Visible = True
        Next I
        For I = linPi.Count To NPI - 2
            Load linPi(I)
            linPi(I).Visible = True
        Next I
    ElseIf NPI < shpPi.Count Then
        For I = shpPi.Count - 1 To NPI Step -1
            Unload shpPi(I)
        Next I
        For I = linPi.Count - 1 To NPI - 1 Step -1
            Unload linPi(I)
        Next I
    End If
'
    With GrillePoint
        .Rows = NPI + 1
        For I = 0 To NPI - 1
            .Row = I + 1
            .Col = 0
            .Text = I
            .Col = 1
            .Text = Format$(Pi(I).X, "#0.000")
            .Col = 2
            .Text = Format$(Pi(I).Y, "#0.000")
            .Col = 3
            .Text = Format$(Pi(I).Z, "#0.000")
        Next I
    End With
'
'
'
End Sub

Private Sub ChargerPoint(ByVal FileNome$)
'
'   Legge, da File, le coordinate dei Point
'   da interpolare:
'
    Dim FF%, X1!, Y1!, X2!, Y2!, Z1!, Z2!, M$
    Dim Pi_T() As P_Type
'
    On Error GoTo ChargerPoint_ERR
'
    X1 = 1E+38
    X2 = -1E+38
    Y1 = 1E+38
    Y2 = -1E+38
    Z1 = 1E+38
    Z2 = -1E+38
    
    NPI = 0
'
    FF = FreeFile
    Open FileNome For Input As #FF
'
    Do
        ReDim Preserve Pi_T(0 To NPI)
        Input #FF, Pi_T(NPI).X, Pi_T(NPI).Y, Pi_T(NPI).Z
'
        If X1 > Pi_T(NPI).X Then X1 = Pi_T(NPI).X
        If X2 < Pi_T(NPI).X Then X2 = Pi_T(NPI).X
        If Y1 > Pi_T(NPI).Y Then Y1 = Pi_T(NPI).Y
        If Y2 < Pi_T(NPI).Y Then Y2 = Pi_T(NPI).Y
        If Z1 > Pi_T(NPI).Z Then Z1 = Pi_T(NPI).Z
        If Z2 < Pi_T(NPI).Z Then Z2 = Pi_T(NPI).Z
        
        NPI = NPI + 1
    Loop While Not EOF(FF)
'
'
ChargerPoint_ERR:
    Close #FF
'
    If Err = 0 Then
        Pi() = Pi_T()
'
        Xmin = Int(X1)
        Xmax = Arrond(X2)
        Ymin = Int(Y1)
        Ymax = Arrond(Y2)
        Zmin = Int(Z1)
        Zmax = Arrond(Z2)
        
'
        ' Dimmensioner le graphique:

        Quadrillage picQuadrillage, Xmin, Xmax, Ymin, Ymax, , , 2, ShOx, ShOy, , "x", "y"
'
        ' Dessiner les Points:
        AjouterPoint
        PositionnerPoint
'
        txtNPI = NPI
        txtXmin = Xmin
        txtXmax = Xmax
        txtYmin = Ymin
        txtYmax = Ymax
        txtZmin = Zmin
        txtZmax = Zmax
        
        ' Affichage en vue XY
        OptionVue(0).Value = True

        
'
    Else
        M$ = "Fichier de point corrompu" & vbNewLine
        M$ = M$ & "Erreur " & Str$(Err.Number) & vbNewLine
        M$ = M$ & Err.Description
        MsgBox M$, vbCritical, "Erreur dans ChargerPoint " & Err.Source
    End If
'
'
'
End Sub

Private Sub DessinCourbe(ByVal PicB As PictureBox)
'
'   Calcul le tracé de la courbe
'   avec les points NPC:
'
    Dim I&
'
    ReDim Pc(0 To NPC - 1) ' As P_Type
'
    ' Calcul de la courbe:
    Select Case TypeC$
        Case "Bezier"
        Call Bezier(Pi(), Pc())
'
        Case "B-Spline"
        Call B_Spline(Pi(), NK, Pc())
'
        Case "C-Spline"
        Call C_Spline(Pi(), Pc())
'
        Case "T-Spline"
        Call T_Spline(Pi(), VZ, Pc())
    End Select
'
    ' Dessin de la courbe:
    PicB.DrawWidth = 1
    
    
    Select Case Vue
    Case 0 'Vue XY
            PicB.PSet (Pc(0).X, Pc(0).Y), vbBlack
            For I = 1 To NPC - 1
                PicB.Line -(Pc(I).X, Pc(I).Y), vbBlack
            Next I
        '
            If chkSegnaPc = vbChecked Then
                ' Affiche les points d'approxiamtion de la spline:
                PicB.DrawWidth = 3
                For I = 0 To NPC - 1
                    PicB.PSet (Pc(I).X, Pc(I).Y), vbMagenta
                Next I
            End If
    
    Case 1 ' Vue XZ
            PicB.PSet (Pc(0).X, Pc(0).Z), vbBlack
            For I = 1 To NPC - 1
                PicB.Line -(Pc(I).X, Pc(I).Z), vbBlack
            Next I
        '
            If chkSegnaPc = vbChecked Then
                ' Affiche les points d'approxiamtion de la spline:
                PicB.DrawWidth = 3
                For I = 0 To NPC - 1
                    PicB.PSet (Pc(I).X, Pc(I).Z), vbMagenta
                Next I
            End If
    
    Case 2 ' Vue YZ
            PicB.PSet (Pc(0).Y, Pc(0).Z), vbBlack
            For I = 1 To NPC - 1
                PicB.Line -(Pc(I).Y, Pc(I).Z), vbBlack
            Next I
        '
            If chkSegnaPc = vbChecked Then
                ' Affiche les points d'approxiamtion de la spline:
                PicB.DrawWidth = 3
                For I = 0 To NPC - 1
                    PicB.PSet (Pc(I).Y, Pc(I).Z), vbMagenta
                Next I
            End If
    End Select
    
'
'
'
End Sub
