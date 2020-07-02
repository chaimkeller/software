VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm GDMDIform 
   BackColor       =   &H8000000C&
   Caption         =   "Map Digitizer"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   -2040
   ClientWidth     =   15300
   Icon            =   "GDBMDIfm.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin VB.Timer GPS_timer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   10080
      Top             =   1080
   End
   Begin VB.Timer Timer_bubble 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   9480
      Top             =   1080
   End
   Begin VB.Timer GPSCom_timer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   8760
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7380
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer CenterPointTimer 
      Enabled         =   0   'False
      Left            =   6420
      Top             =   1080
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   6795
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20223
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Screen coordinates (pixels)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "Magnification"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Object.ToolTipText     =   "Edit type"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4020
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   65
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":1682
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":1796
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":18AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":19BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":1E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":2326
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":2778
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":2BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":3020
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":3474
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":38C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":3D1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":469C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":47F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":4954
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":4AB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":4C0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":4D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":4EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":5318
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":576C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":5BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":5D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":5E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":5FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":6130
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":6584
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":66E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":683C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":699C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":6AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":6C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":70AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":74FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":DD5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":DEB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":E30A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":E75C
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":EBAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":ED08
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":F15A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":F2B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":F40E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":F99D
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":FCEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":FDA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":101FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":10355
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":106A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":10801
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":1095B
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":10AB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":10C0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":10F61
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":1139A
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":11615
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":1176F
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":118C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":123DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":124ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GDBMDIfm.frx":125FF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   15270
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   15300
      Begin MSComCtl2.Animation ani_prg 
         Height          =   200
         Left            =   12900
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   344
         _Version        =   393216
         FullWidth       =   13
         FullHeight      =   13
      End
      Begin VB.ComboBox combContour 
         Height          =   315
         ItemData        =   "GDBMDIfm.frx":12711
         Left            =   11960
         List            =   "GDBMDIfm.frx":1273C
         TabIndex        =   21
         ToolTipText     =   "Contour Interval (meters)"
         Top             =   30
         Visible         =   0   'False
         Width           =   715
      End
      Begin VB.PictureBox picProgBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   12820
         ScaleHeight     =   270
         ScaleWidth      =   2265
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         ToolTipText     =   " X coordinate of cursor"
         Top             =   60
         Width           =   1240
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         ToolTipText     =   " Y coordinate of cursor"
         Top             =   60
         Width           =   1240
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Elevation (meters) at cursor postion"
         Top             =   60
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "Distance (km)"
         Top             =   60
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   9240
         TabIndex        =   5
         Text            =   "0"
         ToolTipText     =   "Y coordinate of center mark"
         Top             =   25
         Width           =   1240
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         ToolTipText     =   "Elevation (m) at center mark (double click to change)"
         Top             =   25
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   7320
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "X coordinate of center mark"
         Top             =   25
         Width           =   1240
      End
      Begin VB.Label Label1 
         Caption         =   "XPix"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   60
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "YPix"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   60
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "height"
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   4000
         TabIndex        =   14
         ToolTipText     =   "meters"
         Top             =   60
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "dist:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5520
         TabIndex        =   13
         ToolTipText     =   "distance (meters)"
         Top             =   60
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "XPix"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         TabIndex        =   12
         Top             =   60
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "YPix"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8760
         TabIndex        =   11
         Top             =   60
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "height"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   10580
         TabIndex        =   10
         ToolTipText     =   "meters"
         Top             =   60
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   55
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MapParameters"
            Object.ToolTipText     =   "Options dialog"
            ImageIndex      =   37
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MapInput"
            Object.ToolTipText     =   "Load default map"
            ImageIndex      =   30
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "MapTopo"
            Object.ToolTipText     =   "contour plot a topo_pixel.xyz file"
            ImageIndex      =   46
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "East"
            Object.ToolTipText     =   "Move to the east (topo maps only)"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "North"
            Object.ToolTipText     =   "Move to the north (topo maps only)"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "West"
            Object.ToolTipText     =   "Move to the west (topo maps only)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "South"
            Object.ToolTipText     =   "Move to the south (topo maps only)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "GoInput"
            Object.ToolTipText     =   "Goto Inputed coordinates"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "GeoCoord"
            Object.ToolTipText     =   "Show geographic coordinates (topo maps only)"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "PrintMap"
            Object.ToolTipText     =   "Print current map frame"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Access"
            Object.ToolTipText     =   "Input data using MSAccess GSI database program"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "EditScannedDB"
            Object.ToolTipText     =   "Edit the scanned database"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
            Object.Width           =   200
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "SearchKey"
            Object.ToolTipText     =   "Activate search mode"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Wizard"
            Object.ToolTipText     =   "Search wizard"
            ImageIndex      =   36
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "MapRetrieve"
            Object.ToolTipText     =   "Coordinates boundaries of search"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "AllSources"
            Object.ToolTipText     =   "Search over all sample sources"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Wells"
            Object.ToolTipText     =   "Search samples from wells"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Outcroppings"
            Object.ToolTipText     =   "Search samples from surface outcroppings"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Fossils"
            Object.ToolTipText     =   "Pick fossil type to search for"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Formations"
            Object.ToolTipText     =   "Pick formation to search for"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Clients"
            Object.ToolTipText     =   "Pick clients, analysts, companies, etc. to search for"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Dates"
            Object.ToolTipText     =   "Pick geological age dates to search"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
            Object.Width           =   200
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Report"
            Object.ToolTipText     =   "Preview database record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "PrintResults"
            Object.ToolTipText     =   "Print search results"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "SaveResults"
            Object.ToolTipText     =   "Save search results"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
            Object.Width           =   200
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ArcMap"
            Object.ToolTipText     =   "Enable digitizing using GTCO table works"
            ImageIndex      =   52
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "GoogleEarth"
            Object.ToolTipText     =   "Export search results to Google Earth"
            ImageIndex      =   38
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "GPSbut"
            Object.ToolTipText     =   "Use GPS to position map"
            ImageIndex      =   39
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Magbut"
            ImageIndex      =   45
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Digitizerbut"
            Object.ToolTipText     =   "Activate Digitizer"
            ImageIndex      =   40
            Object.Width           =   1200
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ExtendGridbut"
            Object.ToolTipText     =   "Extend disappearing grid"
            ImageIndex      =   44
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "RubberrSheetingbut"
            Object.ToolTipText     =   "Rubber Sheeting"
            ImageIndex      =   41
         EndProperty
         BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Eraserbut"
            Object.ToolTipText     =   "Erase Digitized Points"
            ImageIndex      =   43
         EndProperty
         BeginProperty Button41 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Sweepbut"
            Object.ToolTipText     =   "Choose region to erase"
            ImageIndex      =   51
         EndProperty
         BeginProperty Button42 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "EditPointsbut"
            Object.ToolTipText     =   "Edit digitized point positions"
            ImageIndex      =   53
         EndProperty
         BeginProperty Button43 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Hardybut"
            Object.ToolTipText     =   "Calculate Hardy quadratic surfaces"
            ImageIndex      =   42
         EndProperty
         BeginProperty Button44 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "OpenXYZfilebut"
            Object.ToolTipText     =   "contour plot a topo_pixel.xyz file"
            ImageIndex      =   46
         EndProperty
         BeginProperty Button45 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CreateDTMbut"
            ImageIndex      =   64
         EndProperty
         BeginProperty Button46 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Smoothbut"
            Object.ToolTipText     =   "Smooth region of the merged basis DTM"
            ImageIndex      =   65
         EndProperty
         BeginProperty Button47 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button48 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "TableWorksbut"
            Object.ToolTipText     =   "Enable table works"
            ImageIndex      =   52
         EndProperty
         BeginProperty Button49 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
            Object.Width           =   600
         EndProperty
         BeginProperty Button50 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "HeightSearchbut"
            Object.ToolTipText     =   "Search for highest point"
            ImageIndex      =   59
         EndProperty
         BeginProperty Button51 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Contourbut"
            Object.ToolTipText     =   "Generate Contours"
            ImageIndex      =   61
         EndProperty
         BeginProperty Button52 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Profilebut"
            Object.ToolTipText     =   "View horizon profile"
            ImageIndex      =   62
         EndProperty
         BeginProperty Button53 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button54 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button55 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Helpkey"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   6
            Object.Width           =   700
         EndProperty
      EndProperty
      Enabled         =   0   'False
      Begin MSComctlLib.Slider SliderContour 
         Height          =   300
         Left            =   12000
         TabIndex        =   19
         ToolTipText     =   "Sensitivity 70"
         Top             =   0
         Visible         =   0   'False
         Width           =   2200
         _ExtentX        =   3889
         _ExtentY        =   529
         _Version        =   393216
         LargeChange     =   10
         Max             =   140
         SelStart        =   70
         Value           =   70
      End
      Begin VB.CommandButton cmdCancelSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10200
         Picture         =   "GDBMDIfm.frx":12770
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   20
         Visible         =   0   'False
         Width           =   455
      End
      Begin MSComctlLib.ProgressBar prbSearch 
         Height          =   285
         Left            =   10680
         TabIndex        =   18
         Top             =   20
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         Max             =   1000
         Scrolling       =   1
      End
   End
   Begin VB.Menu mnuMaps 
      Caption         =   "&Files"
      Begin VB.Menu mnuOpenSaved 
         Caption         =   "&Open saved results"
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu mnuspacer5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Paths/Options"
      End
      Begin VB.Menu mnuspacer3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSearchOption 
         Caption         =   "&Search Options"
         Visible         =   0   'False
         Begin VB.Menu mnuReportVisible 
            Caption         =   "&Report visible during search"
         End
         Begin VB.Menu mnuReportInvisible 
            Caption         =   "&Minimize report during search"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuAnimation 
         Caption         =   "&Map Animation"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGeo 
         Caption         =   "&Geoids"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuGeoidClark 
            Caption         =   "&Clarke 1880"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuGeoidWGS84 
            Caption         =   "&WGS84 (GPS)"
         End
      End
      Begin VB.Menu mnuspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLocations 
         Caption         =   "&Locations"
         Enabled         =   0   'False
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCoordinateList 
         Caption         =   "&Coordinate List"
      End
      Begin VB.Menu mnuSpace10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintMap 
         Caption         =   "&Print Map"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu retrieveinfofm 
      Caption         =   "&Display"
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
      Begin VB.Menu mnuReport 
         Caption         =   "&Report"
         Enabled         =   0   'False
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPrintReport 
         Caption         =   "&Print Report"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSave 
         Caption         =   "S&ave"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuspace5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcGIS 
         Caption         =   "Arc&GIS"
         Enabled         =   0   'False
         Shortcut        =   ^V
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoogle 
         Caption         =   "Google &Earth"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuDigitize 
      Caption         =   "&Digitizing"
      Enabled         =   0   'False
      Begin VB.Menu mnuDigiExtendGrid 
         Caption         =   "Extend disappearing grid line"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizeRubberSheeting 
         Caption         =   "Grid coordinates"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSeparatorRubberSheeting 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizePoint 
         Caption         =   "&Points (start w/ blank elev.)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizePointSameHeights 
         Caption         =   "&Points (start w/ last elev.)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigiDeleteLastPoint 
         Caption         =   "&Delete last point"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizeDeletPoint 
         Caption         =   "&Delete nearest digitized point"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizeEndPoint 
         Caption         =   "&End point digitizing"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPointSeparator 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizeLine 
         Caption         =   "&Begin line digitizing"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizeEndLine 
         Caption         =   "&End line digitizing"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigiDeleteLastLine 
         Caption         =   "&Delete Last Line"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizeDeleteLine 
         Caption         =   "&Delete Nearest Line"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLineSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDigitizeContour 
         Caption         =   "&Begin countour digitizing"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizeEndContour 
         Caption         =   "&End contour digitizing"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizeSpacer1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEraser 
         Caption         =   "Erase"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigiSweep 
         Caption         =   "Erase rectangular region"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu spacerEraser 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDigitizeHardy 
         Caption         =   "&Hardy quadrac surfaces"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu helpfm 
      Caption         =   "&Help"
      Begin VB.Menu readmefm 
         Caption         =   "&Help Topics on Map Digitizer"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuspacer7 
         Caption         =   "-"
      End
      Begin VB.Menu aboutfm 
         Caption         =   "&About the Map Digitizer"
      End
   End
End
Attribute VB_Name = "GDMDIform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TT1 As New CBalloonToolTip                             '//On Demand tooltip
Dim TT2 As New CBalloonToolTip                             '//mouse over tooltip

Private Sub aboutfm_Click()
   GDfrmAbout.Visible = True
End Sub

Private Sub CenterPointTimer_Timer()

   If Not DigitizeOn And Not DigitizerEraser Then 'DigitizeContour Then
      Call DrawPlotMark(0, 0, 1)
   Else
      Exit Sub
      End If
End Sub

Private Sub cmdCancelSearch_Click()
   StopSearch = True
End Sub

Private Sub DigiTimer_Timer()

'    If maginit Then
'       Call GDMagform.Form_Load
'       Exit Sub
'       End If
'
'    GetCursorPos mouse_digi                                  ' capture mouse-position
'    'Me_digi.Caption = "X: " & mouse.x & ", Y: " & mouse.y    ' write position 2 window-title
'    GDMagform.Text2 = mouse_digi.x
'    GDMagform.Text4 = mouse_digi.y
'    w_digi = GDMagform.Picture2.ScaleWidth                   ' destination width
'    h_digi = GDMagform.Picture2.ScaleHeight                         ' destination height
'    sw_digi = w_digi * (1 / zoom_digi) ' source width
'    sh_digi = h_digi * (1 / zoom_digi) ' source height
'    x_digi = mouse_digi.x * twipsx - sw_digi \ 2                                ' x source position (center to destination)
'    y_digi = mouse_digi.y * twipsy - sh_digi \ 2                                ' y source position (center to destination)
'    GDMagform.Picture2.Cls                              ' clean picturebox
'    StretchBlt GDMagform.Picture2.hDC, 0, 0, w_digi, h_digi, dhdc_digi, x_digi, y_digi, sw_digi, sh_digi, SRCCOPY_digi  ' copy desktop (source) and strech to picturebox (destination)
'
'    centerx = w / 2 '100 'PictureBox1.Width / 2
'    centery = h / 2 '100 'PictureBox1.Height / 2
'    sizecrosshair = w / 50 '50 'PictureBox1.Width / 2
'    a& = GDMagform.Picture2.DrawMode
'    GDMagform.Picture2.DrawMode = 1
'    GDMagform.Picture2.DrawStyle = 0
'    GDMagform.Picture2.DrawWidth = 2
'    GDMagform.Picture2.Line (centerx - sizecrosshair, centery)-(centerx + sizecrosshair, centery), QBColor(12)
'    GDMagform.Picture2.Line (centerx, centery - sizecrosshair)-(centerx, centery + sizecrosshair), QBColor(12)
'    GDMagform.Picture2.DrawMode = a&
End Sub


Private Sub MDIForm_Load()

  '--------------check for other instances of this program--------------
  
   On Error GoTo MDIForm_Load_Error

   If App.PrevInstance = True Then
     'most probably multiple instance of prgram is running
     Call MsgBox("Application already running!", vbExclamation, App.Title)
     PreviousInstance = True
     Call MDIform_QueryUnload(0, 0)
     End If
     
  '--------------------check if running program as administrator------------
   CheckIfAdmin
   
  '-------------------find Windows Version------------------------
   WinVer = GetVista
     
  
  '-------------------SCREEN RESOLUTION CHECK------------------
   'If screen resolution is greater than 1152x864 then give
   'warning.  If resolution less than 800 x 600, then exit.
   Screen.MousePointer = vbHourglass
   XResol = val(Mid$(GetScreenResolution, 1, InStr(GetScreenResolution, "x")))
   Screen.MousePointer = vbDefault
   
'   If XResol > 1152 Then
'      response = MsgBox("The program's map display was designed for a maximum screen" & vbLf & _
'             "resolution of 1152 x 864.  Your resolution exceeds this maximum!" & vbLf & vbLf & _
'             "You should consider changing the screen resolution.  Otherwise the maps won't" & vbLf & _
'             "be displayed properly." & vbLf & vbLf & _
'             "Do you want to exit now and change the resolution?", _
'             vbYesNo + vbExclamation, "MapDigitizer")
'      If response = vbYes Then
'         PreviousInstance = True 'signal to close down without asking
'         Call MDIform_QueryUnload(0, 0)
'         End If
'   ElseIf XResol < 800 Then
'   If XResol < 800 Then
'      MsgBox "To run the MapDigitizer program," & vbLf & _
'             "the screen resolution must be set" & vbLf & _
'             "to a minimum of 800 x 600!", vbOKOnly + vbCritical, "MapDigitizer"
'      End
'      End If
   '-------------------------END CHECK---------------------------
   
   direct$ = App.Path 'This is path that program was installed in.
                       'Teoretically this is the best path, but
                       'in practice it is often not unique.
                       
   'choose default directory for those for those files that
   'must have uniform paths between different computers
   direct2$ = "c:" 'so use such a path
                  'and everybody has a "c" directory
   
   'Nevertheless, check that direct2$ really exists and can be written to
   'Otherwise, step in letters, in case of permission problem, use direct$ instead
   Call CheckDirect2(direct2$)
   
   Digitizing = True 'this is not MapDigitizer
   
   Screen.MousePointer = vbHourglass
   
   GDsplash.Visible = True
   Ret = SetWindowPos(GDsplash.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   waitime = Timer
   Do Until Timer > waitime + 3
      DoEvents
   Loop
   
   Dim RepairInfo As Boolean
   On Error GoTo errdirhand
50
'   'Defaults
'   linked = False
'   heights = False
'   topos = False
'   arcs = False
'   acc = False
   
  '***********kill old temp database files*********
'   errpal& = 1
'   If Dir(direct$ & "\pal_dt_tmp.mdb") <> sEmpty Then
'      Kill direct$ & "\pal_dt_tmp.mdb"
'      End If
'   If Dir(direct$ & "\pal_dt_piv_tmp.mdb") <> sEmpty Then
'      Kill direct$ & "\pal_dt_piv_tmp.mdb"
'      End If
'   errpal& = 0
   
   '**********record in directory information*****
   ReadWriteDefaultsandLink
           
main150:
   
   Screen.MousePointer = vbDefault
  
   '******************************************
   
'   '*********Read tifviewer info*************
'   myfile = Dir(direct$ + "\gdb_tif.sav")
'
'   If myfile = sEmpty Then
'      'see if Windows XP image viewer, shimgvw.dll, exists in the windows/system32 directory
'      tifDir$ = NEDdir
'      tifViewerDir$ = GetSystemPath & "\SHIMGVW.DLL"
'      If Dir(tifViewerDir$) <> sEmpty Then
'         tifCommandLine$ = "RUNDLL32.EXE " & tifViewerDir$ & ", ImageView_Fullscreen"
'         End If
'   Else
'      filin% = FreeFile
'      Open direct$ & "\gdb_tif.sav" For Input As #filin%
'      Line Input #filin%, tifDir$
'      Line Input #filin%, tifViewerDir$
'      Line Input #filin%, tifCommandLine$
'      Close #filin%
'      End If
''*******************************************
 
   Ret = SetWindowPos(GDsplash.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   waitime = Timer
   
   GDMDIform.Visible = True
   
   'get the window's last recorded size and placement
   MeLeft = GetSetting(App.Title, "Settings", "MainLeft", -60)
   MeTop = GetSetting(App.Title, "Settings", "MainTop", -60)
   MeWidth = GetSetting(App.Title, "Settings", "MainWidth", Screen.Width + 120)
   MeHeight = GetSetting(App.Title, "Settings", "MainHeight", Screen.Height - 350)
   
   'never let the program position itself beyond the top of the screen
   If MeTop < 0 Then MeTop = 0
  
   If MeLeft <= 0 And MeTop <= 0 And MeWidth >= Screen.Width And MeHeight >= Screen.Height - 450 Then
      'the last recorded state was bascially the maximized state
      Me.WindowState = vbMaximized
      MeLeft = Me.left
      MeTop = Me.top
   Else
      'leave the window the way it was last closed
      Me.left = MeLeft
      Me.top = MeTop
      Me.Width = MeWidth
      Me.Height = MeHeight
      End If
   
   'disenable buttons and menus until splash form unloads
   GDMDIform.Toolbar1.Enabled = False

   If Installation_Type = 0 Then
'      GDMDIform.Toolbar1.Buttons(47).Visible = False 'version without GTCO digitizer interface
      GDMDIform.Toolbar1.Buttons(32).Visible = False 'version without GTCO digitizer interface
      End If
   
   helpfm.Enabled = False
   
   If Not SplashVis Then GDsplash.Visible = True
   Do Until Timer > waitime + 2
     DoEvents
   Loop
   Unload GDsplash
   Set GDsplash = Nothing
   SplashVis = False
   
   'now enable buttons
   GDMDIform.Toolbar1.Enabled = True
      
   helpfm.Enabled = True
   
   '***********set default timer intervals**************
   'Timer1 is used for map movements (animation)
   'CenterPointTimer is used for blinking the center click mark
   GDMDIform.CenterPointTimer.Interval = 400
   
   MinimizeReport = True 'default state of search report is in
                         'minimized state.  This increases the speed
   
'   'set some more default button states
'   If picnam$ <> sEmpty Then
'      'check if it is really there
'      If Dir(picnam$) = sEmpty Then 'can't find picture
'         GDMDIform.Toolbar1.Buttons(2).Enabled = False 'disenable map buttons
'         GDMDIform.Toolbar1.Buttons(3).Enabled = False 'disenable 1:50000 scale maps
'         GDMDIform.Toolbar1.Buttons(36).Enabled = False
'         End If
'   Else 'no stored picture
'      GDMDIform.Toolbar1.Buttons(2).Enabled = False 'disenable map buttons
'      GDMDIform.Toolbar1.Buttons(3).Enabled = False 'disenable 1:50000 scale maps
'      GDMDIform.Toolbar1.Buttons(36).Enabled = False
'      End If
      
'   'set replace zero ground level with DTM height flags
'   ReplaceWellZ = False
'   ReplaceOtherZ = False
    If UseNewDTM% = 1 Then UsingNewDTM = True
'   If nWellCheck% = 1 Then ReplaceWellZ = True
'   If nOtherCheck% = 1 Then ReplaceOtherZ = True
   
   Screen.MousePointer = vbDefault
   'now give error message if any of the files where not found
   ShowError
   
   'fix coordinates in old database (only need to do once for any one accuracy)
   'so it is commented out
   'FixCoord
           
   If RepairInfo Then 'encountered eof error in loading the paths file
      MsgBox "Errors were encountered while loading the paths and options." & vbLf & _
           "Some defaults were loaded instead." & vbLf & vbLf & _
           "Be sure to check the paths and options with the" & vbLf & _
           "Paths/Options button/menu", vbInformation + vbOKOnly, "MapDigitizer"
       RepairInfo = False
       End If
       
chk5:
   Exit Sub
   
   
errdirhand:
       'something bad is wrong with something else
       'so show error message and ignore all defined paths
        
        If Err.Number = 75 And errpal& = 1 Then
            Screen.MousePointer = vbDefault
            GDsplash.Visible = False
            Unload GDsplash
            SplashVis = False
            errpal& = 0
            'can't kill the old temporary direct$ & "\pal_dt_tmp.mdb directory
            MsgBox "Can't erase the old temporary database: " & vbLf & _
                   direct$ & "\pal_dt_tmp.mdb!" & vbLf & _
                   "Exit the program, erase that file, and start the program again.", _
                   vbOKOnly + vbExclamation, "MapDigitizer"
        ElseIf Err.Number = 62 Then 'old gbinfo.sav file is corrupted--repair it
           'repair as much as possible
           RepairInfo = True
           Resume Next
        Else 'something unexpected
            Screen.MousePointer = vbDefault
            GDsplash.Visible = False
            Unload GDsplash
            SplashVis = False
            Close
            MsgBox "Encountered unexpected error #: " & Err.Number & vbLf & _
                Err.Description & vbLf & vbLf & _
                "Warning: Loading of paths and options was not completed", _
                vbCritical + vbOKOnly, "MapDigitizer"
'            linked = False
'            heights = False
'            topos = False
'            arcs = False
'            acc = False
            Err.Clear
'            GoTo chk5
       
        End If

   On Error GoTo 0
   Exit Sub

MDIForm_Load_Error:

   If Err.Number = 5 Then 'password hasn't been registered yet
      Resume Next
      End If
      
   Screen.MousePointer = vbDefault
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MDIForm_Load of Form GDMDIform"

End Sub

Private Sub MDIform_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo 100
    
    If PreviousInstance Then
       'previous instance of program found
       'so quit without asking
    Else
       response = MsgBox("Do you really want to exit?", vbQuestion + vbYesNoCancel, "MapDigitizer")
       If response <> vbYes Then
          Cancel = True
          Exit Sub
          End If
       End If
    
    'close databases and delete temporary databases
'    If linked Then CloseDatabase
'    If linkedOld Then CloseDatabaseOld
'    If linkedpiv Then CloseDatabasepiv
    
    If Me.WindowState <> vbMinimized Then 'record window's size and position
        SaveSetting App.Title, "Settings", "MainLeft", Me.left
        SaveSetting App.Title, "Settings", "MainTop", Me.top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    If SaveClose% = 1 Then
    
        If infonum& > 0 Then
           Close #infonum&
           infonum& = 0
           End If
    
        infonum& = FreeFile
        Open direct$ + "\gdbinfo.sav" For Output As #infonum&
        Write #infonum&, "This file is used by the MapDigitizer program. Don't erase it!"
        Write #infonum&, dirNewDTM
        Write #infonum&, MinDigiEraserBrushSize
        Write #infonum&, NEDdir
        Write #infonum&, dtmdir
        Write #infonum&, ChainCodeMethod
        Write #infonum&, numDistContour, numDistLines, numSensitivity, numContours ' arcdir, mxddir
        Write #infonum&, PointCenterClick
        Write #infonum&, picnam$
        Write #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
        Write #infonum&, ReportPaths&, DigiSearchRegion, numMaxHighlight&, Save_xyz%
        Write #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
        Write #infonum&, IgnoreAutoRedrawError%
        Write #infonum&, UseNewDTM%, nOtherCheck%
        Write #infonum&, googledir, URL_OutCrop, URL_Well, kmldir, ASTERdir, DTMtype
        Write #infonum&, NX_CALDAT, NY_CALDAT
        Write #infonum&, RSMethod0, RSMethod1, RSMethod2
        Write #infonum&, ULPixX, ULPixY, LRPixX, LRPixY, LRGridX, LRGridY, ULGridX, ULGridY
        Write #infonum&, XStepITM, YStepITM, XStepDTM, YStepDTM, HalfAzi, StepAzi, Apprn, HeightPrecision, DigiConvertToMeters
        Close #infonum&
        End If
    
    'unload forms
    If PicSum Then
       End 'end quickly to avoid waiting for emptying plot buffers
    Else
    
       If GDRSfrmVis Then
          Call WheelUnHook(GDRSfrm.hwnd)
          End If
          
       If DigitizePadVis Then
'          Call WheelUnHook(GDDigitizerfrm.hWnd)
          Unload GDDigitizerfrm
          End If
          
       If DigitizeMagvis Then
          Unload GDDigiMagfrm
          End If
          
       If TabletControlVis Then
          Unload TabConSample_VB_Form
          End If
          
'       If TabletControlOn Then
'          Call CloseTablet
'          End If
    
       'close the forms normally
       Closing = True
       For i& = 1 To Forms.count - 1
          Unload Forms(i&)
       Next i&
          
       Unload GDMDIform
       Set GDMDIform = Nothing
       End If

100   End
End Sub

Private Sub MDIForm_Resize()
   On Error GoTo errhand
   If GeoMap = True Or TopoMap = True Then
        'Repeat the steps taken in GDform1.Form.Load() that
        'position Gdform1, Gdform1.Picture1, and the Scroll Bars
        
        'if the program is minimized, or remaximized avoid this resize
        If GDMDIform.WindowState = vbMinimized Then
           GDform1Height = GDform1.Height 'record height of map form before the minimization -- needed for restoring it
           If DigitizeMagvis Then
              MagWidth = GDDigiMagfrm.Width
              End If
           Exit Sub
           End If
           
        If GDMDIform.ScaleHeight = 0 Then Exit Sub
        
        GDMDIform.top = 0 'don't let it escape off the screen
        
        GDform1.left = 0
        GDform1.top = 0
        
        If Not magvis Then
           GDform1.Width = GDMDIform.ScaleWidth
        Else
           If GDMDIform.WindowState = vbMinimized Then Exit Sub
           '(exit to avoid some nasty resizing!)
           End If
        
        If GDform1.Width > GDMDIform.ScaleWidth Then GDform1.Width = GDMDIform.ScaleWidth
        GDform1.Height = GDMDIform.ScaleHeight
        
        If GDform1Height <> 0 Then GDform1.Height = GDform1Height
        
        If magvis = True Then 'also adjust the size of the magnification window
           GDMagform.top = GDform1.top
           GDMagform.Height = GDform1.Height
           GDform1.Width = GDMagform.left - GDform1.left
           End If
        
        'Initialize location of picture1
        GDform1.Picture1.Move 0, 0, GDform1.ScaleWidth - GDform1.VScroll1.Width, GDform1.ScaleHeight - GDform1.HScroll1.Height
        
        'Position the horizontal scroll bar
        GDform1.HScroll1.top = GDform1.Picture1.Height
        GDform1.HScroll1.left = GDform1.Picture1.left
        GDform1.HScroll1.Width = GDform1.Picture1.Width
        
        'Position the vertical scroll bar
        GDform1.VScroll1.top = 0
        GDform1.VScroll1.left = GDform1.Picture1.Width
        GDform1.VScroll1.Height = GDform1.Picture1.Height
        
        'Set the Max property for the scroll bars.
        GDform1.HScroll1.Max = GDform1.Picture2.Width - GDform1.Picture1.Width
        GDform1.VScroll1.Max = GDform1.Picture2.Height - GDform1.Picture1.Height
            
        'Determine if the child picture will fill up the screen
        'If so, there is no need to use scroll bars.
        GDform1.VScroll1.Visible = (GDform1.Picture1.Height < GDform1.Picture2.Height)
        GDform1.HScroll1.Visible = (GDform1.Picture1.Width < GDform1.Picture2.Width)
        
        'Initiate Scroll Step Sizes
        If GDform1.HScroll1.Visible Then
            GDform1.HScroll1.LargeChange = HScroll1.Max / 20
            GDform1.HScroll1.SmallChange = HScroll1.Max / 60
            End If
        If GDform1.VScroll1.Visible Then
            GDform1.VScroll1.LargeChange = VScroll1.Max / 20
            GDform1.VScroll1.SmallChange = VScroll1.Max / 60
            End If
        
    Else
       GDMDIform.top = MeTop
       End If
       
    Exit Sub
errhand:
    Resume Next
End Sub


Private Sub mnuCoordinateList_Click()
     GDCoordinateList.Visible = True
     BringWindowToTop (GDCoordinateList.hwnd)
End Sub

Public Sub mnuDigiDeleteLastLine_Click()
   'reblit screen setting center at the previous endpoint
   
   'on définit la couleur du pixel courant à partir des pixels alentours
   Dim iBleu As Byte 'stocke la composante bleue à récupèrer
   Dim iVert As Byte 'stocke la composante verte à récupèrer
   Dim iRouge As Byte 'stocke la composante rouge à récupèrer
   
   If numDigiLines = 0 Then Exit Sub
   
   numDigiLines = numDigiLines - 1
   
   UpdateDigiLogFile
   
   ier = ReDrawMap(0)
   If Not InitDigiGraph Then
      InputDigiLogFile 'load up saved digitizing data for the current map sheet
   Else
      ier = RedrawDigiLog
      End If
  
  'now center map on last endpoint
  If numDigiLines > 0 Then Call ShiftMap(CSng(DigiLines(1, numDigiLines - 1).x * DigiZoom.LastZoom), CSng(DigiLines(1, numDigiLines - 1).Y * DigiZoom.LastZoom))
  
  If DigitizePadVis Then
     BringWindowToTop (GDDigitizerfrm.hwnd)
     End If
  
   
End Sub

Public Sub mnuDigiDeleteLastPoint_Click()
   'delete last recorded digitized point
   
   'on définit la couleur du pixel courant à partir des pixels alentours
   Dim iBleu As Byte 'stocke la composante bleue à récupèrer
   Dim iVert As Byte 'stocke la composante verte à récupèrer
   Dim iRouge As Byte 'stocke la composante rouge à récupèrer

   If numDigiPoints = 0 Then Exit Sub
   
   numDigiPoints = numDigiPoints - 1
   
   UpdateDigiLogFile
   
   ier = ReDrawMap(0)
   If Not InitDigiGraph Then
      InputDigiLogFile 'load up saved digitizing data for the current map sheet
   Else
      ier = RedrawDigiLog
      End If
  
  'now center map on last endpoint
  If numDigiPoints > 0 Then Call ShiftMap(CSng(DigiPoints(numDigiPoints - 1).x * DigiZoom.LastZoom), CSng(DigiPoints(numDigiPoints - 1).Y * DigiZoom.LastZoom))
  
  If DigitizePadVis Then
     BringWindowToTop (GDDigitizerfrm.hwnd)
     End If

  
End Sub
Private Sub mnuDigiExtendGrid_Click()

    Dim ier As Integer
    
    If TopoMap Or GeoMap Then
    
        If buttonstate&(38) = 0 Then
           buttonstate&(38) = 1
           GDform1.Picture2.MouseIcon = LoadResPicture(102, vbResCursor) 'load special extend grid cursor
           GDform1.Picture2.MousePointer = vbCustom
           DigitizeExtendGrid = True
           DigiExtendFirstPoint = True
           
'           DigiRS = False
           DigitizePoint = False
           DigitizeBlankPoint = False
           DigitizeLine = False
           DigitizeContour = False
           
            If buttonstate&(37) = 1 Then 'unload digitizing form
               buttonstate&(37) = 0
              GDMDIform.Toolbar1.Buttons(37).value = tbrUnpressed
              DigitizeOn = False
              If DigitizePadVis Then
                 Unload GDDigitizerfrm
                 End If
              End If
              
          'disenable search drags
          If buttonstate&(15) = 1 Then
             buttonstate&(15) = 0
             GDMDIform.Toolbar1.Buttons(15).value = tbrUnpressed
             SearchDigi = False
             End If
             
          'disenable other types of drag window operations
          If DigitizeHardy Then
             DigitizeHardy = False
             buttonstate&(43) = 0
             GDMDIform.Toolbar1.Buttons(43).value = tbrUnpressed
             
            XminC = 0
            YminC = 0
            XmaxC = 0
            YmaxC = 0
             
             End If
          
          If buttonstate&(40) = 1 Then
             buttonstate&(40) = 0
             GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
             DigitizerEraser = False
             GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
             End If
              
          If buttonstate&(41) = 1 Then
             buttonstate&(41) = 0
             GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
             DigitizerSweep = False
             End If
             
          If buttonstate&(42) = 1 Then
             buttonstate&(42) = 0
             GDMDIform.Toolbar1.Buttons(42).value = tbrUnpressed
             DigiEditPoints = False
             End If
           
           'draw all the previously recorded lines
           ier = InputGuideLines
           
           Call MsgBox("Click on the starting and end points of the line" _
                        & vbCrLf & "(Be careful to only use the scroll bars to pan.)" _
                        & vbCrLf & "" _
                        & vbCrLf & "(The order of the points will define the direction of the line.)" _
                        , vbInformation Or vbDefaultButton1, "Extending Grids (Rubber Sheeting)")

        Else
           GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
           buttonstate&(38) = 0
           GDMDIform.Toolbar1.Buttons(38).value = tbrUnpressed
           DigitizeExtendGrid = False
           DigiExtendFirstPoint = False
           ier = ReDrawMap(0) 'erase extended grid lines
           End If
           
        End If
        
   'initialize digitizer mouse coordinates
   digiextendgrid_last.x = INIT_VALUE
   digiextendgrid_last.Y = INIT_VALUE
   digiextendgrid_begin.x = INIT_VALUE
   digiextendgrid_begin.Y = INIT_VALUE
        

End Sub

Public Sub mnuDigiSweep_Click()

    If (TopoMap Or GeoMap) And (numDigiContours > 0 Or numDigiPoints > 0 Or numDigiLines > 0 Or numDigiErase > 0) Then
    
          DigitizePoint = False
          DigitizeLine = False
          DigitizeContour = False
          DigiContourStart = False
          DigitizeHardy = False
          DigiRS = False
          DigitizeExtendGrid = False
          
          If buttonstate&(38) = 1 Then
             buttonstate&(38) = 0
             GDform1.Picture2.MousePointer = vbCrosshair
             DigitizeExtendGrid = False
             DigiExtendFirstPoint = False
             GDMDIform.Toolbar1.Buttons(38).value = tbrUnpressed
             End If
             
          'disenable search drags
          If buttonstate&(15) = 1 Then
             buttonstate&(15) = 0
             GDMDIform.Toolbar1.Buttons(15).value = tbrUnpressed
             SearchDigi = False
             End If
             
          'disenable other types of drag window operations
          If DigitizeHardy Then
             DigitizeHardy = False
             buttonstate&(43) = 0
             GDMDIform.Toolbar1.Buttons(43).value = tbrUnpressed
             
            XminC = 0
            YminC = 0
            XmaxC = 0
            YmaxC = 0
             
             End If
          
          If buttonstate&(40) = 1 Then
             buttonstate&(40) = 0
             GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
             DigitizerEraser = False
             GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
             End If
         
          If buttonstate&(41) = 0 Then
             buttonstate&(41) = 1
             GDMDIform.Toolbar1.Buttons(41).value = tbrPressed
             DigitizerSweep = True
             
             If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(8, 1) 'show the right mode in the digitizer form
            
          Else
             buttonstate&(41) = 0
             GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
             DigitizerSweep = False
             
             If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(8, 0) 'show the right mode in the digitizer form
             
             End If
             
        End If
End Sub

Private Sub mnuDigitize_Click()
    mnuDigitizer_Click
End Sub

Public Sub mnuDigitizeContour_Click()
   Dim ier As Integer
   ier = 0
   
   DigitizeContour = True
   DigitizePoint = False
   DigitizeLine = False
   DigitizeExtendGrid = False
   PointStart = False
   DigiContourStart = True
   DigitizeHardy = False
   DigiRS = False
   
   DigitizerEraser = False
   If buttonstate&(40) = 1 Then
      buttonstate&(40) = 0
      GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
      End If
      
   DigitizerSweep = False
   If buttonstate&(41) = 1 Then
      buttonstate&(41) = 0
      GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
      End If
   
   GDMDIform.SliderContour.Visible = True
   GDMDIform.SliderContour.value = numSensitivity
   
   'redraw all contours and lines with contour color
   ier = ReDrawMap(0)
   ier = RedrawDigiLog
   
   'show magnifier panel if not already activated
   If Not DigitizeMagvis Then
      DigitizeMagInit = True
      GDDigiMagfrm.Visible = True
      GDMDIform.Toolbar1.Buttons(36).value = tbrPressed
      buttonstate&(36) = 1
      End If

   If Not DigitizePadVis Then
      GDDigitizerfrm.Visible = True
      End If
                  
   BringWindowToTop (GDDigitizerfrm.hwnd)
    
   GDDigitizerfrm.txtX = blink_mark.x
   GDDigitizerfrm.txtY = blink_mark.Y
   If Trim$(GDDigitizerfrm.txtelev) = gsEmpty Then
      GDDigitizerfrm.txtelev = Format(str$(ContourHeight / MapUnits), "#####0.0#")
      End If
End Sub

Public Sub mnuDigitizeDeleteLine_Click()

   'delete nearest line to blinking cursor
   
   If numDigiLines = 0 Then Exit Sub
   
   Dim NearX As Single
   Dim NearY As Single

   'on définit la couleur du pixel courant à partir des pixels alentours
   Dim iBleu As Byte 'stocke la composante bleue à récupèrer
   Dim iVert As Byte 'stocke la composante verte à récupèrer
   Dim iRouge As Byte 'stocke la composante rouge à récupèrer

   NearX = blink_mark.x
   NearY = blink_mark.Y

    'find nearest line to the cursor
    dist0 = 999999
    pointnum& = 0
    For i& = 0 To numDigiLines - 1
       dist1 = Sqr((NearX - DigiLines(0, i&).x * DigiZoom.LastZoom) ^ 2 + (NearY - DigiLines(0, i&).Y * DigiZoom.LastZoom) ^ 2)
       dist2 = Sqr((NearX - DigiLines(1, i&).x * DigiZoom.LastZoom) ^ 2 + (NearY - DigiLines(1, i&).Y * DigiZoom.LastZoom) ^ 2)
       Dist = (dist1 + dist2) * 0.5
       If Dist < dist0 Then
           pointnum& = i&
           dist0 = Dist
           End If
    Next i&
     
    'now shift the line array to eliminate this line
    For i& = pointnum& + 1 To numDigiLines - 1
        DigiLines(0, i& - 1).x = DigiLines(0, i&).x
        DigiLines(1, i& - 1).x = DigiLines(1, i&).Y
    Next i&
     
    numDigiLines = numDigiLines - 1
  
  'update the log file
  UpdateDigiLogFile
  
  ier = ReDrawMap(0)
  
  If Not InitDigiGraph Then
     InputDigiLogFile 'load up saved digitizing data for the current map sheet
  Else
     ier = RedrawDigiLog
     End If
  
  'now center map on last endpoint
  If numDigiLines > 0 Then Call ShiftMap(CSng(DigiLines(1, numDigiLines - 1).x * DigiZoom.LastZoom), CSng(DigiLines(1, numDigiLines - 1).Y * DigiZoom.LastZoom))
  
  If DigitizePadVis Then
     BringWindowToTop (GDDigitizerfrm.hwnd)
     End If

End Sub

Public Sub mnuDigitizeDeletPoint_Click()

   'delete point nearest to blinking cursor
   
   Dim NearX As Single
   Dim NearY As Single
   
   'on définit la couleur du pixel courant à partir des pixels alentours
   Dim iBleu As Byte 'stocke la composante bleue à récupèrer
   Dim iVert As Byte 'stocke la composante verte à récupèrer
   Dim iRouge As Byte 'stocke la composante rouge à récupèrer

   If numDigiPoints = 0 Then Exit Sub
   
   NearX = blink_mark.x
   NearY = blink_mark.Y

    'find nearest point to the cursor
    dist0 = 999999
    pointnum& = 0
    For i& = 0 To numDigiPoints - 1
       Dist = Sqr((NearX - DigiPoints(i&).x * DigiZoom.LastZoom) ^ 2 + (NearY - DigiPoints(i&).Y * DigiZoom.LastZoom) ^ 2)
       If Dist < dist0 Then
           pointnum& = i&
           dist0 = Dist
           End If
    Next i&
    
    'now shift the point array to eliminate this point
    For i& = pointnum& + 1 To numDigiPoints - 1
        DigiPoints(i& - 1) = DigiPoints(i&)
    Next i&
     
     numDigiPoints = numDigiPoints - 1
     
  'update log file
  UpdateDigiLogFile
  
  ier = ReDrawMap(0)
  
  If Not InitDigiGraph Then
     InputDigiLogFile 'load up saved digitizing data for the current map sheet
  Else
     ier = RedrawDigiLog
     End If
  
  'now center map on last endpoint
  Call ShiftMap(CSng(DigiPoints(numDigiPoints).x * DigiZoom.LastZoom), CSng(DigiPoints(numDigiPoints).Y * DigiZoom.LastZoom))
  
  If DigitizePadVis Then
     BringWindowToTop (GDDigitizerfrm.hwnd)
     End If
  
End Sub

Public Sub mnuDigitizeEndContour_Click()
   Dim ier As Integer
   ier = 0
   
   DigitizeContour = False
   PointStart = False
   DigiContourStart = False
   
   GDMDIform.SliderContour.Visible = False
   
   'replot the line colors according to height if so flagged
   ier = ReDrawMap(0)
   ier = RedrawDigiLog
   
    If DigitizePadVis And Not DigitizeOn Then
       Unload GDDigitizerfrm
       End If
   
   'start blinker again
    GDMDIform.CenterPointTimer.Enabled = True
    ce& = 1
   
   
End Sub

Public Sub mnuDigitizeEndLine_Click()
    DigitizeEndLine = True
    DigitizeBeginLine = False
    DigitizeLine = False
'    DigitizeDrawLine = False
    If DigitizePadVis And Not DigitizeOn Then
       Unload GDDigitizerfrm
       End If
'   DigitizeLine = False
'
   'erase last line
    gddm = GDform1.Picture2.DrawMode
    gddw = GDform1.Picture2.DrawWidth
    GDform1.Picture2.DrawMode = 7 'erase mode
    GDform1.Picture2.DrawWidth = Max(2, CInt(DigiZoom.LastZoom))
'    GDform1.Picture2.Line (digi_last.X * DigiZoom.LastZoom, digi_last.Y * DigiZoom.LastZoom)-(digi_begin.X * DigiZoom.LastZoom, digi_begin.Y * DigiZoom.LastZoom), QBColor(12)
    GDform1.Picture2.Line (digi_last.x, digi_last.Y)-(digi_begin.x, digi_begin.Y), QBColor(12)
    GDform1.Picture2.DrawMode = gddm
    GDform1.Picture2.DrawWidth = gddw

   digi_last.x = INIT_VALUE
   digi_last.Y = INIT_VALUE
   digi_begin.x = INIT_VALUE
   digi_begin.Y = INIT_VALUE

End Sub

Public Sub mnuDigitizeEndPoint_Click()
    DigitizePoint = False
    DigitizeBlankPoint = False
    If Not DigitizeOn Then Unload GDDigitizerfrm
End Sub

Public Sub mnuDigitizeHardy_Click()

    Dim ier As Integer
    
    If TopoMap Or GeoMap And (DigiRubberSheeting Or RSMethod0) Then
    
        If buttonstate&(43) = 0 Then
        
            XminC = 0
            YminC = 0
            XmaxC = 0
            YmaxC = 0
        
'           If DigitizePadVis And Not DigitizeHardy Then
'              If DigiBackground = &HC0FFC0 Then
'                 'press enter key, then "H" key
'                 KeyDown (vbKeyReturn)
'
'                 KeyDown (vbKeyH)
'                 KeyDown (vbKeyReturn)
'              Else
'                 KeyDown (vbKeyH)
'                 KeyDown (vbKeyReturn)
'                 End If
'              End If
              
           buttonstate&(43) = 1
           GDMDIform.Toolbar1.Buttons(43).value = tbrPressed
           DigitizeHardy = True
            
           If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(13, 1) 'show the right mode in the digitizer form
           
           DigiRS = False
           DigitizePoint = False
           DigitizeBlankPoint = False
           DigitizeLine = False
           DigitizeContour = False
           DigitizeExtendGrid = False
           DigitizerEraser = False
           
           'shut off magnitizer window
           If DigitizeMagvis Then
              Unload GDDigiMagfrm
              End If
              
           If GDRSfrmVis Then
              Unload GDRSfrm
              End If
              
'           DigitizeOn = False
'           If DigitizePadVis Then
'              Unload GDDigitizerfrm
'              End If

'           shut off erasing
            GDMDIform.Toolbar1.Buttons(40).Enabled = False
            GDMDIform.mnuEraser.Enabled = False
            GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
            DigitizerEraser = False
            buttonstate&(40) = 0
           
            GDMDIform.Toolbar1.Buttons(41).Enabled = False
            GDMDIform.mnuDigiSweep.Enabled = False
            GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
            DigitizerSweep = False
            buttonstate&(41) = 0
            
            'shut off point editing
            GDMDIform.Toolbar1.Buttons(42).Enabled = False
            GDMDIform.Toolbar1.Buttons(42).value = tbrUnpressed
            buttonstate&(42) = 0
            
            'shut off searching
            GDMDIform.Toolbar1.Buttons(15).value = tbrUnpressed
            GDMDIform.Toolbar1.Buttons(15).Enabled = False
            buttonstate&(15) = 0
            SearchDigi = False
            
            'shut down height searches and contours
            HeightSearch = False
            GenerateContours = False
'            DigiReDrawContours = False
            GDMDIform.Toolbar1.Buttons(50).value = tbrUnpressed
            GDMDIform.Toolbar1.Buttons(51).value = tbrUnpressed
            buttonstate&(50) = 0
            buttonstate&(51) = 0
            GDMDIform.Toolbar1.Buttons(50).Enabled = False
            GDMDIform.Toolbar1.Buttons(51).Enabled = False
            numContourPoints = 0 'zero contour lines array
            ReDim ContourPoints(numContourPoints)
            ReDim contour(0) 'zero contour color array

            'load previously recorded digitizing results
            ier = ReDrawMap(0)
            If Not InitDigiGraph Then
               InputDigiLogFile 'load up saved digitizing data for the current map sheet
            Else
               ier = RedrawDigiLog
               End If
            
            'shut down blinkers
            ce& = 0 'reset blinker flag
            If GDMDIform.CenterPointTimer.Enabled = True Then
               ce& = 1 'flag that timer has been shut down during drag
               GDMDIform.CenterPointTimer.Enabled = False
               End If
                   
                
            GDMDIform.combContour.Visible = True
            If numContours > 0 Then
               GDMDIform.combContour.Text = str(numContours)
            Else
               GDMDIform.combContour.ListIndex = 6 '10m default '2 3 m default
               End If
           
        Else
        
'           If DigitizePadVis And DigitizeHardy Then
'              If DigiBackground = &HC0FFC0 Then
'                 'press enter key, then "H" key
'                 KeyDown (vbKeyReturn)
'                 End If
'              End If
              
            DigiReDrawContours = False
            numContourPoints = 0 'zero contour lines array
            ReDim ContourPoints(numContourPoints)
            ReDim contour(0) 'zero contour color array
              
           GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
           buttonstate&(43) = 0
           GDMDIform.Toolbar1.Buttons(43).value = tbrUnpressed
           
           XminC = 0
           YminC = 0
           XmaxC = 0
           YmaxC = 0
          
           DigitizeHardy = False
           
           If (heights Or BasisDTMheights) Then
              GDMDIform.Toolbar1.Buttons(50).Enabled = True
              GDMDIform.Toolbar1.Buttons(51).Enabled = True
              End If
           
           If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(13, 0)
           
           'clear the picture of contours and repaint
           ier = ReDrawMap(0)
           
           If DigitizeOn Then
              If Not InitDigiGraph Then
                 InputDigiLogFile 'load up saved digitizing data for the current map sheet
              Else
                 ier = RedrawDigiLog
                 End If
              End If
              
           'disenable making DTM's and profiling
           GDMDIform.Toolbar1.Buttons(45).Enabled = False
           GDMDIform.Toolbar1.Buttons(45).value = tbrUnpressed
           buttonstate&(45) = 0
           DTMcreating = False
           
           GDMDIform.Toolbar1.Buttons(15).Enabled = True
           
           If DigitizePadVis Then
           
              If numDigiContours > 0 Then 'allow for erasing
                 GDMDIform.Toolbar1.Buttons(40).Enabled = True
                 GDMDIform.mnuEraser.Enabled = True
                 GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
                 buttonstate&(40) = 0
                 End If
              
               GDMDIform.Toolbar1.Buttons(41).Enabled = True
               GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
               buttonstate&(41) = 0
            
               GDMDIform.Toolbar1.Buttons(42).Enabled = True
               GDMDIform.Toolbar1.Buttons(42).value = tbrUnpressed
               buttonstate&(42) = 0
               
               DigitizeOn = True
               
               End If
                 
      
           'renable blinking
           GDMDIform.CenterPointTimer.Enabled = True
           ce& = 1
           
           GDMDIform.combContour.Visible = False
           
            XminC = 0
            YminC = 0
            XmaxC = 0
            YmaxC = 0
           
           End If
           
    ElseIf DigiRubberSheeting = False And Not RSMethod0 Then
    
        Call MsgBox("You must first choose a coordinate conversion method!" _
                    & vbCrLf & "" _
                    & vbCrLf & "(Hint: Press the ''Rubber Sheeting  button'')" _
                    , vbInformation, "Hardy Quadratic Surfaces Error")
        
                    
        If DigitizeOn And DigitizePadVis Then
        
           DigiEntered = False
           DigiBackground = &HC0C0FF
           GDDigitizerfrm.lblFunction.BackColor = DigiBackground
           
           End If
        
        End If
    
End Sub

Public Sub mnuDigitizeLine_Click()
'   GDDigitizerfrm.Visible = True
'   BringWindowToTop (GDDigitizerfrm.hWnd)
   DigitizePoint = False
   DigitizeBlankPoint = False
   DigitizeLine = True
   DigitizeBeginLine = True
   DigiRS = False
   DigitizeExtendGrid = False
   DigitizeHardy = False
   DigitizeDeleteContour = False
   
   DigitizerEraser = False
   If buttonstate&(40) = 1 Then
      buttonstate&(40) = 0
      GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
      End If
      
   DigitizerSweep = False
   If buttonstate&(41) = 1 Then
      buttonstate&(41) = 0
      GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
      End If
      
   
   GDMDIform.SliderContour.Visible = False

'   DigitizeDrawLine = False
End Sub

Public Sub mnuDigitizePoint_Click()
   
   DigitizePoint = False
   DigitizeBlankPoint = True
   DigitizeLine = False
   DigitizeContour = False
   DigiRS = False
   DigitizeExtendGrid = False
   DigitizeHardy = False
   DigitizeDeleteContour = False
   
   DigitizerEraser = False
   If buttonstate&(40) = 1 Then
      buttonstate&(40) = 0
      GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
      End If
      
   DigitizerSweep = False
   If buttonstate&(41) = 1 Then
      buttonstate&(41) = 0
      GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
      End If
   
   
   GDMDIform.SliderContour.Visible = False
   
   If Not DigitizePadVis Then
     GDDigitizerfrm.Visible = True
     BringWindowToTop (GDDigitizerfrm.hwnd)
   Else
     BringWindowToTop (GDDigitizerfrm.hwnd)
     End If

End Sub

Public Sub mnuDigitizePointSameHeights_Click()
   
   DigitizePoint = True
   DigitizeBlankPoint = False
   DigitizeLine = False
   DigitizeContour = False
   DigiRS = False
   DigitizeExtendGrid = False
   DigitizeHardy = False
   DigitizeDeleteContour = False
   
   GDMDIform.SliderContour.Visible = False
   
   If Not DigitizePadVis Then
     GDDigitizerfrm.Visible = True
     BringWindowToTop (GDDigitizerfrm.hwnd)
   Else
     BringWindowToTop (GDDigitizerfrm.hwnd)
     End If
   
End Sub

Public Sub mnuDigitizer_Click() '<<<<<<<<<<<<digi changes
    Dim ier As Integer
    
    If buttonstate&(37) = 0 Then
       buttonstate&(37) = 1
       GDMDIform.Toolbar1.Buttons(37).value = tbrPressed
'
'      'shut down blinkers
'      GDMDIform.CenterPointTimer.Enabled = False
'
'      If CenterBlinkState And ce& = 1 Then
'         Call DrawPlotMark(0, 0, 1)
'         End If

        'disenable other types of drag window operations
        If DigitizeHardy Then
           DigitizeHardy = False
           buttonstate&(43) = 0
           GDMDIform.Toolbar1.Buttons(43).value = tbrUnpressed
           
           XminC = 0
           YminC = 0
           XmaxC = 0
           YmaxC = 0
           
           End If
        
        If buttonstate&(41) = 1 Then
          buttonstate&(41) = 0
          GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
          DigitizerSweep = False
          End If
        
        If buttonstate&(40) = 1 Then
          buttonstate&(40) = 0
          GDform1.Picture2.MousePointer = vbCrosshair
          DigitizerEraser = False
          End If
          
        If buttonstate&(38) = 1 Then
           GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
           buttonstate&(38) = 0
           GDMDIform.Toolbar1.Buttons(38).value = tbrUnpressed
           DigitizeExtendGrid = False
           DigiExtendFirstPoint = False
           End If
           
        'shut down height searches and contours
        HeightSearch = False
        GenerateContours = False
        GDMDIform.Toolbar1.Buttons(50).value = tbrUnpressed
        GDMDIform.Toolbar1.Buttons(51).value = tbrUnpressed
        buttonstate&(50) = 0
        buttonstate&(51) = 0
        GDMDIform.Toolbar1.Buttons(50).Enabled = False
        GDMDIform.Toolbar1.Buttons(51).Enabled = False
        
          
      ce& = 0 'reset blinker flag
      If GDMDIform.CenterPointTimer.Enabled = True Then
         ce& = 1 'flag that timer has been shut down during drag
         GDMDIform.CenterPointTimer.Enabled = False
         End If
         
       DigitizeOn = True
'       If Not DigitizeMagvis Then
'          DigitizeMagInit = True
'          GDDigiMagfrm.Visible = True
'          End If
          
       'load previously recorded digitizing results
       ier = ReDrawMap(0)
       If Not InitDigiGraph Then
          InputDigiLogFile 'load up saved digitizing data for the current map sheet
       Else
          ier = RedrawDigiLog
          End If
       
       If DigiRS Then
          Unload GDRSfrm
          End If
       
       If numDigiContours > 0 Then 'allow for erasing
          GDMDIform.Toolbar1.Buttons(40).Enabled = True
          GDMDIform.mnuEraser.Enabled = True
          GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
          buttonstate&(40) = 0
          End If
          
        If numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0 Or numDigiErase > 0 And GDMDIform.mnuDigiSweep.Enabled = False Then
           GDMDIform.Toolbar1.Buttons(41).Enabled = True
           GDMDIform.mnuDigiSweep.Enabled = True
           GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
           buttonstate&(41) = 0
           End If
           
        If numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0 Or numDigiErase > 0 Then 'allow for point editing
           GDMDIform.Toolbar1.Buttons(42).Enabled = True
           GDMDIform.Toolbar1.Buttons(42).value = tbrUnpressed
           buttonstate&(42) = 0
           End If
           
          
         'this in and out call to tracecontours8 fixes some sort of bug
         Call tracecontours8(GDform1.Picture2, INIT_VALUE, 9)
          
         If Not DigitizePadVis Then
            GDDigitizerfrm.Visible = True
            BringWindowToTop (GDDigitizerfrm.hwnd)
         Else
            BringWindowToTop (GDDigitizerfrm.hwnd)
            If DigiRightButtonIndex >= 9 Or DigiRightButtonIndex <= 11 Then
                DigiBackground = &HE0E0E0       'disenabled
                GDDigitizerfrm.lblFunction.Enabled = False
                GDDigitizerfrm.lblFunction.BackColor = DigiBackground
                End If
            End If
       
    Else
       buttonstate&(37) = 0
       GDMDIform.Toolbar1.Buttons(37).value = tbrUnpressed
       
       InitDigiGraph = False
       
       If DigitizePadVis Then Unload GDDigitizerfrm
       
       'always unpress the eraser
       GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
       buttonstate&(40) = 0
       GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
       GDMDIform.Toolbar1.Buttons(40).Enabled = False
       GDMDIform.mnuEraser.Enabled = False
       DigitizerEraser = False
       
       GDMDIform.Toolbar1.Buttons(41).Enabled = False 'disenable sweep erasing
       GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
       GDMDIform.mnuDigiSweep.Enabled = False
       buttonstate&(41) = 0
       
       GDMDIform.Toolbar1.Buttons(42).Enabled = False 'disenable point editing
       GDMDIform.Toolbar1.Buttons(42).value = tbrUnpressed
       buttonstate&(42) = 0
       
      'reenable searches and contours
'      If ((Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm") Or _
'         (Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat")) And _

       If heights And (RSMethod1 Or RSMethod2 Or RSMethod0) And DigiRubberSheeting Then
         GDMDIform.Toolbar1.Buttons(50).Enabled = True
         GDMDIform.Toolbar1.Buttons(51).Enabled = True
         End If
       
       'refresh map
       DigitizeOn = False
       ier = ReDrawMap(0)
       
       'renable blinking
       GDMDIform.CenterPointTimer.Enabled = True
       ce& = 1
      
       End If
       
   'initialize digitizer mouse coordinates
   digi_last.x = INIT_VALUE
   digi_last.Y = INIT_VALUE
   digi_begin.x = INIT_VALUE
   digi_begin.Y = INIT_VALUE
       
End Sub

Public Sub mnuDigitizeRubberSheeting_Click()

  Dim ier As Integer
  
  If buttonstate&(39) = 0 Then
     buttonstate&(39) = 1
     
    If DigitizeHardy Then
       DigitizeHardy = False
       buttonstate&(43) = 0
       GDMDIform.Toolbar1.Buttons(43).value = tbrUnpressed
       
      XminC = 0
      YminC = 0
      XmaxC = 0
      YmaxC = 0
       
       End If
     
     
     If DigiRubberSheeting And Not RSMethod0 Then
        Select Case MsgBox("Rubber sheeting seems to be complete for this map!" _
                           & vbCrLf & "" _
                           & vbCrLf & "Do you want to redo it?" _
                           , vbYesNo Or vbQuestion Or vbDefaultButton2, "Rubber Sheeting")
        
           Case vbYes
              'load grid intersections and plot them
              If Not DigiRubberSheeting Then
              
                  ier = ReDrawMap(0)
                  If ier = 0 Then
                     ier = InputGuideLines 'plot guide lines if any
                     If ier = 0 Then
                        ier = ReadRSfile 'plot grid extension lines
                        End If
                     End If
                     
                  If ier <> 0 Then 'error detected
                     buttonstate&(39) = 0
                     Exit Sub
                     End If
                     
                  End If
              GDRSfrm.Visible = True
              BringWindowToTop (GDRSfrm.hwnd)
              buttonstate&(39) = 0
              GDMDIform.Toolbar1.Buttons(39).value = tbrUnpressed
              
           Case vbNo
              'close button
              buttonstate&(39) = 0
              GDMDIform.Toolbar1.Buttons(39).value = tbrUnpressed
              DigiRS = False
              
            If RSopenedfile Then 'close the log file
               Close RSfilnum%
               RSopenedfile = False
               RSfilnum% = 0
               End If
               
            'reload map without the x's
            ier = ReDrawMap(0)
            
        End Select
     'now load any RS file and plot it
     ElseIf Not RSMethod0 Then
        'now load any RS file and plot it
        DigiRS = True
        
        DigitizePoint = False
        DigitizeBlankPoint = False
        DigitizeLine = False
        DigitizeContour = False
        DigitizeExtendGrid = False
        
        'load grid intersections and plot them
        If Not DigiRubberSheeting Then
           ier = ReDrawMap(0)
           If ier = 0 Then
              ier = InputGuideLines 'plot guide lines if any
              If ier = 0 Then
                 ier = ReadRSfile 'plot grid extension lines
                 End If
              End If
               
           If ier <> 0 Then 'error detected
              buttonstate&(39) = 0
              Exit Sub
              End If
           
           GDRSfrm.Visible = True
           BringWindowToTop (GDRSfrm.hwnd)

           If numRS = NX_CALDAT * NY_CALDAT And NX_CALDAT <> 0 And NY_CALDAT <> 0 Then
        
                Select Case MsgBox("Rubber sheeting seems to be complete for this map!" _
                                   & vbCrLf & "" _
                                   & vbCrLf & "(Hint: Press the ''Activate Calculation Method'' button to activate it.)" _
                                   , vbOKOnly + vbInformation, "Rubber Sheeting")
                
                  Case vbOK
                     'the rubber sheeting is complete, so run it
                     GDRSfrm.cmdConvert.Enabled = True
                     Close #RSfilnum%  'no need to keep it open
                     RSopenedfile = False
                     buttonstate&(39) = 0
                     GDMDIform.Toolbar1.Buttons(39).value = tbrUnpressed
                     
                End Select
                
                End If
           
           End If
    ElseIf RSMethod0 Then
        'do nothing
        GDRSfrm.Visible = True
        BringWindowToTop (GDRSfrm.hwnd)
        GDRSfrm.Height = 1000 '7365
        GDRSfrm.chkSimple.value = vbChecked
        GDRSfrm.chkRS.value = vbUnchecked
        If NX_CALDAT = 0 Or NY_CALDAT = 0 Then
           GDRSfrm.chkRS.Enabled = False
           End If
           
        If ((Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm") Or _
            (Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat")) Then
            'enable GPS button
            GDMDIform.Toolbar1.Buttons(34).Enabled = True
            End If
            
        If (heights Or BasisDTMheights) And (RSMethod1 Or RSMethod2 Or RSMethod0) And DigiRubberSheeting Then

            'enable search height button
            GDMDIform.Toolbar1.Buttons(50).Enabled = True
            'enable contour generation
            GDMDIform.Toolbar1.Buttons(51).Enabled = True
            
            GDMDIform.Label1 = lblX
            GDMDIform.Label5 = lblX
            GDMDIform.Label2 = LblY
            GDMDIform.Label6 = LblY
            
            GDMDIform.Text3.Visible = True
            GDMDIform.Label3.Visible = True
            GDMDIform.Text7.Visible = True
            GDMDIform.Label7.Visible = True
            
            GDMDIform.Text4.Visible = True
            GDMDIform.Label4.Visible = True
                
            End If
        
        End If
    
  ElseIf buttonstate&(39) = 1 Then
  
     If DigitizeHardy And buttonstate&(43) = 1 Then
        Select Case MsgBox("Warning: You are about to close coordinate readoutss" _
                           & vbCrLf & "which will also disenable Hardy Quadratic Surface Analysis" _
                           & vbCrLf & "as well as Height Searching and Contouring." _
                           & vbCrLf & "" _
                           & vbCrLf & "Proceed?" _
                           & vbCrLf & "" _
                           , vbYesNo Or vbInformation Or vbDefaultButton2, "Rubber Sheeting")
        
          Case vbYes
        
          Case vbNo
            Exit Sub
        End Select
        
        End If
        
     buttonstate&(39) = 0
     GDMDIform.Toolbar1.Buttons(39).value = tbrUnpressed
     
    'disenable GPS, map button
    GDMDIform.Toolbar1.Buttons(34).Enabled = False
    GDMDIform.Toolbar1.Buttons(50).Enabled = False
    GDMDIform.Toolbar1.Buttons(34).value = tbrUnpressed
    GDMDIform.Toolbar1.Buttons(50).value = tbrUnpressed
    GDMDIform.Toolbar1.Buttons(51).value = tbrUnpressed
    buttonstate&(34) = 0
    buttonstate&(50) = 0
    buttonstate&(51) = 0
    
    GDMDIform.Label1 = "XPix"
    GDMDIform.Label5 = "XPix"
    GDMDIform.Label2 = "YPix"
    GDMDIform.Label6 = "YPix"
    
    GDMDIform.Text3.Visible = False
    GDMDIform.Label3.Visible = False
    GDMDIform.Text7.Visible = False
    GDMDIform.Label7.Visible = False
            
    GDMDIform.Text4.Visible = False
    GDMDIform.Label4.Visible = False
     
     
     DigiRS = False
     If GDRSfrmVis Then
        Unload GDRSfrm
        End If
     
     If RSopenedfile Then 'close the log file
        Close RSfilnum%
        RSopenedfile = False
        RSfilnum% = 0
        End If
        
     If DigitizeHardy And buttonstate&(43) = 1 Then
        'disenable hardy selections
        mnuDigitizeHardy_Click
        End If
        
      'disenable Hardy quadratic surface analaysis
      GDMDIform.Toolbar1.Buttons(43).Enabled = False
      'disenable search height button
      GDMDIform.Toolbar1.Buttons(50).Enabled = False
      'disenable contour button
      GDMDIform.Toolbar1.Buttons(51).Enabled = False
        
     If DigiRubberSheeting Then
        DigiRubberSheeting = False
   
        'renable blinking
        GDMDIform.CenterPointTimer.Enabled = True
        ce& = 1
        End If
        
     End If


End Sub

Public Sub mnuEraser_Click()
    If (TopoMap Or GeoMap) And numDigiContours > 0 Then
    
          DigitizePoint = False
          DigitizeLine = False
          DigitizeContour = False
          DigiContourStart = False
          DigitizeHardy = False
          DigiRS = False
          DigitizeExtendGrid = False
          
          If buttonstate&(38) = 1 Then
             buttonstate&(38) = 0
             GDform1.Picture2.MousePointer = vbCrosshair
             DigitizeExtendGrid = False
             DigiExtendFirstPoint = False
             GDMDIform.Toolbar1.Buttons(38).value = tbrUnpressed
             End If
          
          If buttonstate&(41) = 1 Then
             buttonstate&(41) = 0
             GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
             DigitizerSweep = False
             End If
         
          If buttonstate&(40) = 0 Then
             buttonstate&(40) = 1
             GDMDIform.Toolbar1.Buttons(40).value = tbrPressed
             GDform1.Picture2.MouseIcon = LoadResPicture(101, vbResCursor) 'load special eraser cursor
             GDform1.Picture2.MousePointer = vbCustom
             DigitizerEraser = True
             
             If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(7, 1) 'show the right mode in the digitizer form
            
          Else
             GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
             buttonstate&(40) = 0
             GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
             DigitizerEraser = False
             
             If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(7, 0) 'show the right mode in the digitizer form
             
             End If
        
        End If
    
End Sub

Private Sub mnuGoogle_Click()

        If Searching Then
           MsgBox "Can't perform this operation in the middle of a search!", _
                  vbExclamation + vbOKOnly, "MapDigitizer"
           Exit Sub
           End If
        
        If Not SearchDigi Then
           MsgBox "To use this option, you must define a search area!" & vbCrLf & _
                  "To define searches press the Search button and drag" & vbCrLf & _
                  "over the area of interest.", vbExclamation Or vbOKOnly, "MapDigitizer"
           Exit Sub
           End If
        
        'Write search results to external database
        'and then activate ArcGIS.
        GoogleDump = True
        If google = False Then
           response = MsgBox("Sorry, you haven't yet defined the path for Google Earth" & vbLf & _
                  "Use the Option Menu to help find a path." & vbLf & _
                  "Do you still wish to create a KLM file?" & vbLf & _
                  "(The KLM file will be called c:\GSI_Search_(date).)", _
                   vbExclamation + vbYesNoCancel, "MapDigitizer")
           If response = vbYes Then
           
              If PicSum Then 'report available, ask if want to copy it
                 
                 Select Case MsgBox("A search report is visible." _
                                 & vbCrLf & "Do you wish to export the search results to Google Earth?" _
                                 , vbYesNo + vbQuestion + vbDefaultButton1, App.Title)
               
                    Case vbYes
                       
                       Call ExportReportToGoogleEarth(KMLFileName$, 0, ier%)
                       
                       If ier% = -1 Then 'error detected, abort
                          Exit Sub
                          
                       ElseIf ier% < -1 Then
                          Call MsgBox("The Illegal character, ""&"", was found in search report #: " & Trim$(str$(-ier% - 1)) & "'s place name." _
                                    & vbCrLf & "Where ever the Illegal character was found it was replaced with a ""+""" _
                                    , vbInformation, "Export to Google Earth")
                          GoTo mG500
                          
                       Else 'succeeded in the copy
                          GoTo mG500
                          End If
                    
                    Case Else
                        Exit Sub 'user aborted
                              
                 End Select
                 
                 End If
                 
              End If
              
           GoogleDump = False
           Exit Sub
        Else
           
           If PicSum Then 'report available, ask if want to copy it
               
              Select Case MsgBox("A search report is visible." _
                                 & vbCrLf & "Do you wish to export the Name, ITMx, ITMy, Z fields to Google Earth?" _
                                 , vbYesNo + vbQuestion + vbDefaultButton1, App.Title)
               
                 Case vbYes
                    Call ExportReportToGoogleEarth(KMLFileName$, 0, ier%)
                    
                    If ier% = -1 Then 'error detected, abort
                       Exit Sub
                       
                    ElseIf ier% < -1 Then
                       Call MsgBox("The Illegal character, ""&"", was found in search report #: " & Trim$(str$(-ier% - 1)) & "'s place name." _
                                 & vbCrLf & "Where ever the Illegal character was found it was replaced with a ""+""" _
                                 , vbInformation, "Export to Google Earth")
                       GoTo mG500
                       
                    Else 'succeeded in the copy
                       GoTo mG500
                       End If
                       
                 Case Else
                    Exit Sub 'user aborted
                 
              End Select
              
              End If
              
          End If
           
mG500:  defkml$ = InputBox("The name of the Google Earth KLM file generated is: " & KMLFileName$ & vbLf & _
                     "You can enter a different file name if you wish." & vbLf & vbLf & _
                     "Press OK to use the default name.", "Google kml file", KMLFileName$)
        If defkml$ = sEmpty Then 'user canceled
           GoogleDump = False
           Exit Sub
        Else
           If defkml$ <> KMLFileName$ Then
              'check if already exists
              If Dir(defkml$) <> sEmpty Then
                  Select Case MsgBox("The kml file name you picked already exists!" _
                                     & vbCrLf & "Do you want to overwrite the existing file?" _
                                     , vbYesNoCancel Or vbQuestion Or vbDefaultButton2, "Google Earth kml filename")
                  
                   Case vbYes
                       FileCopy KMLFileName$, defkml$
                  
                   Case vbNo, vbCancel
                       GoTo mG500
                       
                  End Select
              
              Else
                 FileCopy KMLFileName$, defkml$
                 End If
              
              End If
           End If
      
        num = Shell(googledir & "\googleearth.exe " & Chr$(34) & defkml$ & Chr$(34), vbMaximizedFocus)
        GoogleDump = False

End Sub



Private Sub mnuExit_Click()
   Unload Me
   Set GDMDIform = Nothing
End Sub


Private Sub mnuGeoidClark_Click()
   If mnuGeoidClark.Checked = True Then
      mnuGeoidClark.Checked = False
      mnuGeoidWGS84.Checked = True
      GpsCorrection = True
   Else
      mnuGeoidWGS84.Checked = False
      mnuGeoidClark.Checked = True
      GpsCorrection = False
      End If
      
   If Geo Then 'geographic coordinate converter visible, so display geoid info
      If GpsCorrection Then
         GDGeoFrm.Caption = "Geographic Coordinates" & " - WGS84 geoid"
      Else
         GDGeoFrm.Caption = "Geographic Coordinates" & " - Clark geoid"
         End If
      End If
      
End Sub

Private Sub mnuGeoidWGS84_Click()
   If mnuGeoidWGS84.Checked = True Then
      mnuGeoidWGS84.Checked = False
      mnuGeoidClark.Checked = True
      GpsCorrection = False
   Else
      mnuGeoidWGS84.Checked = True
      mnuGeoidClark.Checked = False
      GpsCorrection = True
      End If
      
   If Geo Then 'geographic coordinate converter visible, so display geoid info
      If GpsCorrection Then
         GDGeoFrm.Caption = "Geographic Coordinates" & " - WGS84 geoid"
      Else
         GDGeoFrm.Caption = "Geographic Coordinates" & " - Clark geoid"
         End If
      End If
      
End Sub



Private Sub mnuHardy_Click()
   AA = 1
End Sub

'Private Sub mnuGotoRetrieve_Click()
'   Call gotocoord
'End Sub

Private Sub mnuLocations_Click()
    If SearchVis Then
       Ret = ShowWindow(GDSearchfrm.hwnd, SW_MINIMIZE)
       End If
    GDLocationfrm.Visible = True
    BringWindowToTop (GDLocationfrm.hwnd)
End Sub


Private Sub mnuOpenSaved_Click()
   'look for saved search results and reload them into gdreportfrm
   On Error GoTo errhand
   
10 GDMDIform.CommonDialog1.FileName = sEmpty
   GDMDIform.CommonDialog1.Filter = "Comma separated text (*.txt)|*.txt"
   GDMDIform.CommonDialog1.FilterIndex = 1
   GDMDIform.CommonDialog1.ShowOpen
   'check for existing files, and for wrong save directories
  
   If GDMDIform.CommonDialog1.FileName = sEmpty Then Exit Sub
   
   ext$ = RTrim$(Mid$(GDMDIform.CommonDialog1.FileName, InStr(1, GDMDIform.CommonDialog1.FileName, ".") + 1, 3))
   If ext$ <> "txt" Or Dir(GDMDIform.CommonDialog1.FileName) = sEmpty Then
      MsgBox "You must select a file with the extension ""txt""!", vbExclamation + vbOKOnly, "MapDigitizer"
      Exit Sub
      End If
          
   filrpt& = FreeFile
   Open GDMDIform.CommonDialog1.FileName For Input As #filrpt&
   
   Line Input #filrpt&, doclin$
   'check for correct file
   If InStr(doclin$, "MapDigitizer Search Results, Date/Time: ") = 0 Then
      response = MsgBox("The file you requested is not a listing of search results!" & vbLf & _
             "Please request a different file.", vbExclamation + vbOKCancel, "MapDigitizer")
      If response = vbOK Then
         GoTo 10
      Else
         Exit Sub
         End If
      End If
   
   Screen.MousePointer = vbHourglass
   
   'prepare report's list view
   If PicSum Then
     'clear old items if any, also clear plot buffer
      GDReportfrm.cmdClear = True
      GDReportfrm.lvwReport.ListItems.Clear
      GDReportfrm.lvwReport.ColumnHeaders.Clear
      GDMDIform.StatusBar1.Panels(2) = sEmpty 'erase any nonapplicable status message
      
      'make sure that report is visible during search
      If Not MinimizeReport Then
         'ret = BringWindowToTop(GDReportfrm.hwnd)
         Ret = ShowWindow(GDReportfrm.hwnd, SW_MAXIMIZE)
         End If
   Else 'load it
      PicSum = True
      GDReportfrm.Visible = True
      End If
             
   Line Input #filrpt&, doclin$
   Line Input #filrpt&, doclin$
   'read in number of columns and rows
   Input #filrpt&, NumCol&, numRow&
   Line Input #filrpt&, doclin$
   Line Input #filrpt&, doclin$
  
  'read in and load up column headers
   For i& = 1 To NumCol&
     Input #filrpt&, doclin$
     GDReportfrm.lvwReport.ColumnHeaders.Add , , doclin$, 1500
     GDReportfrm.lvwReport.ColumnHeaders(i&).Alignment = lvwColumnLeft
   Next i&
   
   Line Input #filrpt&, doclin$
   Line Input #filrpt&, doclin$
   
   'read in and load up search results
   For j& = 1 To numRow&
      For i& = 1 To NumCol&
         If i& = 1 Then
            Input #filrpt&, doclin$
            Set mitem = GDReportfrm.lvwReport.ListItems.Add()
            mitem.Text = Trim$(doclin$)
         ElseIf i& <> 4 Then
            Input #filrpt&, doclin$
            mitem.SubItems(i& - 1) = Trim$(doclin$)
         ElseIf i& = 4 Then 'fossil information
            Input #filrpt&, doclin$
            mitem.SubItems(i& - 1) = Trim$(doclin$)
            
            'determine the icon by what fossils are present
            IconNum& = 0
            If InStr(doclin$, "cono") Then
               IconNum& = 1
               End If
            
            If InStr(doclin$, "diatom") Then
               If IconNum& <> 0 Then
                  IconNum& = -1
                  GoTo 50
               Else
                  IconNum& = 2
                  End If
               End If
            
            If InStr(doclin$, "foram") Then
               If IconNum& <> 0 Then
                  IconNum& = -1
                  GoTo 50
               Else
                  IconNum& = 3
                  End If
               End If
            
            If InStr(doclin$, "mega") Then
               If IconNum& <> 0 Then
                  IconNum& = -1
                  GoTo 50
               Else
                  IconNum& = 4
                  End If
               End If
            
            If InStr(doclin$, "nan") Then
               If IconNum& <> 0 Then
                  IconNum& = -1
                  GoTo 50
               Else
                  IconNum& = 5
                  End If
               End If
            
            If InStr(doclin$, "ostra") Then
               If IconNum& <> 0 Then
                  IconNum& = -1
                  GoTo 50
               Else
                  IconNum& = 6
                  End If
               End If
            
            If InStr(doclin$, "palyn") Then
               If IconNum& <> 0 Then
                  IconNum& = -1
                  GoTo 50
               Else
                  IconNum& = 7
                  End If
               End If
            
50            Select Case IconNum&
              Case -1
                 mitem.SmallIcon = "multi"
              Case 0
                 mitem.SmallIcon = "blank"
              Case 1
                 mitem.SmallIcon = "cono"
              Case 2
                 mitem.SmallIcon = "diatom"
              Case 3
                 mitem.SmallIcon = "foram"
              Case 4
                 mitem.SmallIcon = "mega"
              Case 5
                 mitem.SmallIcon = "nano"
              Case 6
                 mitem.SmallIcon = "ostra"
              Case 7
                 mitem.SmallIcon = "paly"
              Case Else
            End Select

            
            End If

      Next i&
   Next j&

   Close #filtm1&
   'unselect first record (no records will be highlighted)
   GDReportfrm.lvwReport.ListItems.Item(1).Selected = False
   numReport& = numRow&
   
   GDReportfrm.txtZmin = "0" 'initialize filter heights
   GDReportfrm.txtZmax = "0"

   Screen.MousePointer = vbDefault
   
   Exit Sub
   
errhand:
   Screen.MousePointer = vbDefault
   MsgBox "Encountered error #: " & Err.Number & vbLf & _
          Err.Description & vbLf & _
          sEmpty, vbCritical + vbOKOnly, "MapDigitizer"
 
End Sub

Private Sub mnuOptions_Click()
  'the button state of this form is controlled by the Activate and
  'Deactivate events, and the load and unload events
  If OptionsVis Then
     Unload GDOptionsfrm
'     BringWindowToTop (GDOptionsfrm.hWnd)
  Else
     GDOptionsfrm.Visible = True
     BringWindowToTop (GDOptionsfrm.hwnd)
     End If
End Sub



Private Sub mnuPrintMap_Click()
    
    If magvis Then 'magnification window is visible
       MsgBox "When the magnification window is visible, you can" & vbLf & _
              "only print the portion of the map that is magnified" & vbLf & vbLf & _
              "To print the portion that has been magnified," & vbLf & _
              "use the print button in the magnification window", _
              vbInformation + vbOKOnly, "MapDigitizer"
       Exit Sub
       End If
       
    If Previewing Then 'show the print preview
       BringWindowToTop (PrintPreview.hwnd)
       MsgBox "If you want to print the map, first close this preview!", vbInformation + vbOKOnly, "MapDigitizer"
       Exit Sub
       End If

    'dump the map screen to the Print Preview for printing
    If (TopoMap Or GeoMap) And Not ScreenDump Then 'And Not PicSum And Not SearchDB Then
       'print current map sheet
       BringWindowToTop (GDform1.hwnd) 'make sure map is on top of z-order
       If PicSum Or SearchDigi Then 'wait a bit for map to repaint
         waitime = Timer
         Do Until Timer > waitime + 0.5
            DoEvents
         Loop
       End If
       ScreenDump = True
       PrintPreview.Visible = True
    ElseIf (TopoMap Or GeoMap) And ScreenDump And Not PicSum And Not SearchDigi Then
       'make sure its still visible
       'ret = BringWindowToTop(PrintPreview.hwnd)
       Ret = ShowWindow(PrintPreview.hwnd, SW_MAXIMIZE)
    'ElseIf PicSum Then
    '   MsgBox "Can't print the maps as long as a search report is visible!", _
    '          vbExclamation + vbOKOnly, "MapDigitizer"
    'ElseIf SearchDB Then
    '   MsgBox "Can't print the maps as long as searches are activated!", _
    '          vbExclamation + vbOKOnly, "MapDigitizer"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuPrintReport_Click
' DateTime  : 10/1/2002 18:33
' Purpose   : Print Preview a database record
'---------------------------------------------------------------------------------------
'
Private Sub mnuPrintReport_Click()
        
         If PreviewOrderNum& = 0 And NewHighlighted& = 0 Then
            'haven't picked any record to preview
            If PicSum Then
               MsgBox "You haven't picked any record for previewing!" & vbLf & vbLf & _
                      "To preview search results either:" & vbLf & _
                      "(1) Double click on the desired record." & vbLf & _
                      "(2) Click on the desired record and press the Print button.", _
                      vbExclamation + vbOKOnly, "MapDigitizer"
            Else
               MsgBox "You haven't specified which record to preview!", _
                      vbExclamation + vbOKOnly, "MapDigitizer"
               End If
            Exit Sub
            End If
         
         'if printpreview already visible, restore it to top Z order
         If Previewing Then
            Ret = ShowWindow(PrintPreview.hwnd, SW_MAXIMIZE)
            Exit Sub
            End If
        
        'check if report is available
        If Not PicSum Then
           'clarify if want to preview order number
'           If Not SearchDB Then Call PreviewOrder
        Else
           PrintPreview.Visible = True
           Ret = ShowWindow(PrintPreview.hwnd, SW_MAXIMIZE)
           End If
End Sub

Public Sub mnuReport_Click()

   Dim ier As Integer
   Dim XGeo As Double, YGeo As Double

        If (Not SearchDigi And (numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0 Or numDigiErase > 0)) Or _
            (DigitizeOn And (numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0 Or numDigiErase > 0)) Then
           'load digitizer points and lines that were entered
           
            GDMDIform.Toolbar1.Buttons(28).Enabled = True
           
            With GDReportfrm
            
                 'clear the List View of old results
                .lvwReport.ListItems.Clear
                .lvwReport.ColumnHeaders.Clear
                .lvwReport.Sorted = False
                .lvwReport.Refresh
            
               .lvwReport.ColumnHeaders.Add , , "Type", 2000
               .lvwReport.ColumnHeaders.Add , , "Index", 1000
               .lvwReport.ColumnHeaders.Add , , "X coord", 1000
               .lvwReport.ColumnHeaders.Add , , "Y coord", 1000
               If DigiRubberSheeting Then
                  'add geographic coordinates also
                  .lvwReport.ColumnHeaders.Add , , lblX, 3000
                  .lvwReport.ColumnHeaders.Add , , LblY, 3000
                  End If
               .lvwReport.ColumnHeaders.Add , , "elevation", 1000
               
               .lvwReport.ColumnHeaders(1).Alignment = lvwColumnLeft
               .lvwReport.ColumnHeaders(2).Alignment = lvwColumnLeft
               .lvwReport.ColumnHeaders(3).Alignment = lvwColumnLeft
               .lvwReport.ColumnHeaders(4).Alignment = lvwColumnLeft
               .lvwReport.ColumnHeaders(5).Alignment = lvwColumnLeft
               If DigiRubberSheeting Then
                  .lvwReport.ColumnHeaders(6).Alignment = lvwColumnLeft
                  .lvwReport.ColumnHeaders(7).Alignment = lvwColumnLeft
                  End If
               
               'now add the points vertices
               For i& = 0 To numDigiPoints - 1
                   
                   If SearchDigi And _
                      (DigiPoints(i&).x < ReportCoord(0).x Or DigiPoints(i&).x > ReportCoord(1).x Or _
                       DigiPoints(i&).Y < ReportCoord(0).Y Or DigiPoints(i&).Y > ReportCoord(1).Y) Then
                   Else
                   
                        Set mitem = GDReportfrm.lvwReport.ListItems.Add()
                        mitem.Text = "point"
                        mitem.SmallIcon = "point2"
                   
                        If Not DigiRubberSheeting Then
                           mitem.SubItems(1) = str$(i&)
                           mitem.SubItems(2) = str$(DigiPoints(i&).x) 'X
                           mitem.SubItems(3) = str$(DigiPoints(i&).Y) 'Y
                           mitem.SubItems(4) = str$(DigiPoints(i&).Z) 'Z
                        Else
                           mitem.SubItems(1) = str$(i&)
                           mitem.SubItems(2) = str$(DigiPoints(i&).x) 'X
                           mitem.SubItems(3) = str$(DigiPoints(i&).Y) 'Y
                           GoSub ScreenToGeo
                           mitem.SubItems(4) = str$(XGeo)
                           mitem.SubItems(5) = str$(YGeo)
                           mitem.SubItems(6) = str$(DigiPoints(i&).Z) 'Z
                           End If
                    End If
                      
               Next i&
               
               'now add lines vertices
               For i& = 0 To numDigiLines - 1
                   
                   If SearchDigi And _
                      (DigiLines(0, i&).x < ReportCoord(0).x Or DigiLines(0, i&).x > ReportCoord(1).x Or _
                       DigiLines(0, i&).Y < ReportCoord(0).Y Or DigiLines(0, i&).Y > ReportCoord(1).Y Or _
                       DigiLines(1, i&).x < ReportCoord(0).x Or DigiLines(1, i&).x > ReportCoord(1).x Or _
                       DigiLines(1, i&).Y < ReportCoord(0).Y Or DigiLines(1, i&).Y > ReportCoord(1).Y) Then

                   Else
                     
                        Set mitem = GDReportfrm.lvwReport.ListItems.Add()
                        mitem.Text = "line vertex 1"
                        mitem.SmallIcon = "line"
                        
                        If Not DigiRubberSheeting Then
                           mitem.SubItems(1) = str$(i&)
                           mitem.SubItems(2) = str$(DigiLines(0, i&).x) 'X
                           mitem.SubItems(3) = str$(DigiLines(0, i&).Y) 'Y
                           mitem.SubItems(4) = str$(DigiLines(0, i&).Z) 'Z
                        Else
                           mitem.SubItems(1) = str$(i&)
                           mitem.SubItems(2) = str$(DigiLines(0, i&).x) 'X
                           mitem.SubItems(3) = str$(DigiLines(0, i&).Y) 'Y
                           GoSub ScreenToGeo
                           mitem.SubItems(4) = str$(XGeo)
                           mitem.SubItems(5) = str$(YGeo)
                           mitem.SubItems(6) = str$(DigiLines(0, i&).Z) 'Z
                           End If
                   
                        Set mitem = GDReportfrm.lvwReport.ListItems.Add()
                        mitem.Text = "line vertex 2"
                        mitem.SmallIcon = "line"
                     
                        If Not DigiRubberSheeting Then
                           mitem.SubItems(1) = str$(i&)
                           mitem.SubItems(2) = str$(DigiLines(1, i&).x) 'X
                           mitem.SubItems(3) = str$(DigiLines(1, i&).Y) 'Y
                           mitem.SubItems(4) = str$(DigiLines(1, i&).Z) 'Z
                        Else
                           mitem.SubItems(1) = str$(i&)
                           mitem.SubItems(2) = str$(DigiLines(1, i&).x) 'X
                           mitem.SubItems(3) = str$(DigiLines(1, i&).Y) 'Y
                           GoSub ScreenToGeo
                           mitem.SubItems(4) = str$(XGeo)
                           mitem.SubItems(5) = str$(YGeo)
                           mitem.SubItems(6) = str$(DigiLines(1, i&).Z) 'Z
                           End If
                           
                        End If
                        
                Next i&
               
               'now add contour vertices
                For i& = 0 To numDigiContours - 1
                
                   If SearchDigi And _
                      (DigiContours(i&).x < ReportCoord(0).x Or DigiContours(i&).x > ReportCoord(1).x Or _
                       DigiContours(i&).Y < ReportCoord(0).Y Or DigiContours(i&).Y > ReportCoord(1).Y) Then
                   Else
                
                        Set mitem = GDReportfrm.lvwReport.ListItems.Add()
                        mitem.Text = "Contour point"
                        mitem.SmallIcon = "Contour"
                          
                        If Not DigiRubberSheeting Then
                           mitem.SubItems(1) = str$(i&)
                           mitem.SubItems(2) = str$(DigiContours(i&).x) 'X
                           mitem.SubItems(3) = str$(DigiContours(i&).Y) 'Y
                           mitem.SubItems(4) = str$(DigiContours(i&).Z) 'Z
                        Else
                           mitem.SubItems(1) = str$(i&)
                           mitem.SubItems(2) = str$(DigiContours(i&).x) 'X
                           mitem.SubItems(3) = str$(DigiContours(i&).Y) 'Y
                           GoSub ScreenToGeo
                           mitem.SubItems(4) = str$(XGeo)
                           mitem.SubItems(5) = str$(YGeo)
                           mitem.SubItems(6) = str$(DigiContours(i&).Z) 'Z
                           End If
                           
                        End If
                      
               Next i&
               
               'now add erased points
                For i& = 0 To numDigiErase - 1
                
                   If SearchDigi And _
                      (DigiErasePoints(i&).x < ReportCoord(0).x Or DigiErasePoints(i&).x > ReportCoord(1).x Or _
                       DigiErasePoints(i&).Y < ReportCoord(0).Y Or DigiErasePoints(i&).Y > ReportCoord(1).Y) Then
                   Else
                
                        Set mitem = GDReportfrm.lvwReport.ListItems.Add()
                        mitem.Text = "Erased point"
                        mitem.SmallIcon = "Eraser"
                          
                        mitem.SubItems(1) = str$(i&)
                        mitem.SubItems(2) = str$(DigiErasePoints(i&).x) 'X
                        mitem.SubItems(3) = str$(DigiErasePoints(i&).Y) 'Y
                        mitem.SubItems(4) = "0"
                        
                      End If
                      
               Next i&
               
            
           End With
           
           Ret = ShowWindow(GDReportfrm.hwnd, SW_MAXIMIZE)
           GDMDIform.StatusBar1.Panels(1) = sEmpty 'erase any nonapplicable status message
           
           numReport& = GDReportfrm.lvwReport.ListItems.count
           
           Exit Sub
           End If
           
         numReport& = GDReportfrm.lvwReport.ListItems.count
         NumReportPnts& = numReport&

         If PicSum = True Then
            'form is already in use but hidden
            'so make it reappear
            Ret = ShowWindow(GDReportfrm.hwnd, SW_MAXIMIZE)
            GDMDIform.StatusBar1.Panels(1) = sEmpty 'erase any nonapplicable status message
         Else
           'clarify if want to preview order number
'            Call PreviewOrder
            End If
            
    Exit Sub
    
ScreenToGeo:
  'convert screen coordinates to geographic coordinates
    If RSMethod1 Then
       ier = RS_pixel_to_coord2(CDbl(val(mitem.SubItems(1))), CDbl(val(mitem.SubItems(2))), XGeo, YGeo)
    ElseIf RSMethod2 Then
       ier = RS_pixel_to_coord(CDbl(val(mitem.SubItems(1))), CDbl(val(mitem.SubItems(2))), XGeo, YGeo)
    ElseIf RSMethod0 Then
       ier = Simple_pixel_to_coord(CDbl(val(mitem.SubItems(1))), CDbl(val(mitem.SubItems(2))), XGeo, YGeo)
       End If
Return
End Sub

Private Sub mnuReportInvisible_Click()
   If mnuReportInvisible.Checked = True Then
      mnuReportInvisible.Checked = False
      mnuReportVisible.Checked = True
      MinimizeReport = False
   ElseIf mnuReportInvisible.Checked = False Then
      mnuReportInvisible.Checked = True
      mnuReportVisible.Checked = False
      MinimizeReport = True
      End If
End Sub

Private Sub mnuReportVisible_Click()
   If mnuReportVisible.Checked = True Then
      mnuReportVisible.Checked = False
      mnuReportInvisible.Checked = True
      MinimizeReport = True
   ElseIf mnuReportVisible.Checked = False Then
      mnuReportVisible.Checked = True
      mnuReportInvisible.Checked = False
      MinimizeReport = False
      End If
End Sub

'help files
Private Sub readmefm_Click()

On Error GoTo ErrHandler


If Dir(direct$ & "\MapDigitizer.hlp") <> sEmpty Then
    Shell ("Winhlp32.exe " & direct$ & "\MapDigitizer.hlp")
'   GDMDIform.CommonDialog1.CancelError = True
'   GDMDIform.CommonDialog1.HelpCommand = cdlHelpContents
'   GDMDIform.CommonDialog1.HelpFile = direct$ & "\MapDigitizer.hlp"
'   GDMDIform.CommonDialog1.ShowHelp
Else
   If Dir(direct$ & "\Map Digitizer Program.pdf") <> sEmpty Then
      SuperShell direct$ & "\Map Digitizer Program.pdf", , SA_Open, SW_SHOWNORMAL, , , True
   Else
      MsgBox "The help file: " & direct$ & "\MapDigitizer.hlp" & " not found!"
      End If
   End If
   
ErrHandler:
    ' User pressed Cancel button.
    Exit Sub
        
End Sub

Private Sub SliderContour_Click()
   SliderContour.ToolTipText = "Sensitivity: " & str(SliderContour.value)
End Sub

Private Sub Text7_DblClick()
   NewHeight = InputBox("Enter the new elevation (meters):", "Elevation change", GDMDIform.Text7.Text)
   If IsNumeric(val(NewHeight)) Then
      GDMDIform.Text7.Text = NewHeight
   Else
      MsgBox "Please enter a number!", vbInformation + vbOKOnly, "Input Error"
      End If
End Sub

Private Sub Timer_bubble_Timer()
   TT1.Destroy
   Timer_bubble.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   'handles button click events from GDMDIform toolbar

'   On Error GoTo Toolbar1_ButtonClick_Error

    GDMDIform.StatusBar1.Panels(1) = sEmpty
    GDMDIform.StatusBar1.Panels(2) = sEmpty

   Select Case Button.Key
     Case "MapInput" 'load or unload Geo maps (left Geo map button)
           GeoMapMode% = 0
           mnuMapInput_Click
     Case "MapParameters" 'map parameters
           mnuOptions_Click
     Case "GoInput" 'move on map to coordinates
           Call gotocoord
     Case "PrintMap"
           'dump the map screen to the Print Preview for printing
           mnuPrintMap_Click
     Case "Report"
          mnuReport_Click
     Case "PrintResults"
          mnuPrintReport_Click
     Case "SaveResults"
          mnuSave_Click
     Case "SearchKey"
          mnuSearchActivated_Click
     Case "GoogleEarth"
          mnuGoogle_Click
     Case "GPSbut"
          mnuGPS_Click
     Case "Magbut"
          Call DigiMag
     Case "Digitizerbut" '<<<<<<<<<<<<digi changes
          mnuDigitizer_Click
     Case "ExtendGridbut"
          mnuDigiExtendGrid_Click
     Case "RubberrSheetingbut"
          mnuDigitizeRubberSheeting_Click
     Case "Eraserbut"
          mnuEraser_Click
     Case "Sweepbut"
          mnuDigiSweep_Click
     Case "EditPointsbut"
          Call EditDigitizedPoints
     Case "Hardybut"
          mnuDigitizeHardy_Click
     Case "MapTopo", "OpenXYZfilebut"
          Call OpenXYZfile
     Case "CreateDTMbut"
          Call generateDTM
     Case "TableWorksbut", "ArcMap"
          'enable tableworks with tableworks mouse capture
          Call mnuTablet
'          TabConSample_VB_Form.Visible = True
     Case "Smoothbut"
          mnuSmooth_Click
     Case "HeightSearchbut"
          mnuSearchHeights_Click
     Case "Contourbut"
          mnuContour_Click
     Case "Profilebut"
          mnuProfile_Click
     Case "Helpkey"
          readmefm_Click
     Case Else
   End Select
   
   'now redepress or undepress the buttons, and refresh
   For i% = 1 To Toolbar1.Buttons.count
      If buttonstate&(i%) = 1 Then
         GDMDIform.Toolbar1.Buttons(i%).value = tbrPressed
      Else
         GDMDIform.Toolbar1.Buttons(i%).value = tbrUnpressed
         End If
   Next i%
   GDMDIform.Toolbar1.Refresh 'refresh the visual state of the toolbar
   
   Screen.MousePointer = vbDefault

   On Error GoTo 0
   Exit Sub

Toolbar1_ButtonClick_Error:

    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Toolbar1_ButtonClick of Form GDMDIform"
End Sub
Sub mnuMapInput_Click()
  'loads and unloads the Geo Maps
  
    If magvis Then 'close magnify window first
       MsgBox "Can't close or switch maps until you close the magnification window!", _
              vbOKOnly + vbExclamation, "MapDigitizer"
       Exit Sub
       End If
  
    mode% = GeoMapMode%
    
    DigiZoom.LastZoom = 1# 'beginning zoom is 100%
10:
    With GDMDIform
  
    If (buttonstate&(2) = 0 And mode% = 0) Or (buttonstate&(18) = 0 And mode% = 1) Then
50:     myfile = Dir(picnam$)
        If myfile = sEmpty Then 'try using the current directory
           picnamtrial$ = App.Path & "\" & picnam$
           If Dir(picnatrial$) <> sEmpty Then picnam$ = picnamtrial$
           GoTo 50
           End If
           
        If myfile = sEmpty Then
           response = MsgBox("Can't find map!" & vbLf & _
                      "Use the Files: Paths/Options menu to help find it.", _
                      vbCritical + vbOKOnly, "GSIDB")
           'take further response
           GeoMap = False
           Exit Sub
        Else
           Screen.MousePointer = vbHourglass
           If buttonstate&(3) = 1 Then 'topo maps button pressed, reset it
              buttonstate&(3) = 0
              .Toolbar1.Buttons(3).value = tbrUnpressed
              For i& = 4 To 7
                .Toolbar1.Buttons(i&).Enabled = False
              Next i&
              .Toolbar1.Buttons(9).Enabled = False
              .Toolbar1.Refresh
              If Geo = True Then 'GeoCoord visible, so unload it
                 Geo = False
                 Unload GDGeoFrm
                 Set GDGeoFrm = Nothing
                 .Toolbar1.Buttons(9).value = tbrUnpressed
                 buttonstate&(0) = 1
                 End If
              End If
           If buttonstate&(15) = 1 Then 'search still activated
              .Toolbar1.Buttons(15).value = tbrPressed
              End If
              
           .mnuGeo.Enabled = False 'disenable menu of geo. coordinates display
           .Toolbar1.Buttons(2).value = tbrPressed
            buttonstate&(2) = 1
           
'           If topos = True Then .Toolbar1.Buttons(3).Enabled = True
'           .Toolbar1.Buttons(8).Enabled = True 'enable goto coordinates
'           .Toolbar1.Buttons(10).Enabled = True 'enable print maps
'           .mnuPrintMap.Enabled = True
'           .Label1 = lblX
'           .Label5 = lblX
'           .Label2 = LblY
'           .Label6 = LblY
          
           If (Mid$(LCase(lblX), 1, 3) <> "itm" Or Mid$(LCase(LblY), 1, 3) <> "itm") And Not Digitizing Then
             
             If GeoX <> 0 Or GeoY <> 0 Then
               'put map in old place
               MsgBox "Geo map coordinate system was not ITM!" & vbLf & _
                      "Geo map will be placed in its last position.", vbExclamation + vbOKOnly, "MapDigitizer"
               Call UpdatePositionFile(GeoX * DigiZoom.LastZoom, GeoY * DigiZoom.LastZoom, GeoHgt)
             
             Else
                MsgBox "Geo map coordinate system was not ITM!" & vbLf & _
                    "Geo map will be placed in middle of map.", vbExclamation + vbOKOnly, "MapDigitizer"
                       
                'shift the map to center of map and record position
                GeoX = x10 + (x20 - x10) / 2
                GeoY = y20 + (y10 - y20) / 2
                GeoHgt = 0
                Call UpdatePositionFile(GeoX * DigiZoom.LastZoom, GeoY * DigiZoom.LastZoom, GeoHgt)
                End If
             End If
           .Toolbar1.Refresh
           
           'load up Geo map
           Call ShowGeoMap(0)
           If g_ier = -1 Then 'shutting down gracefully from out of memory error
              g_ier = 0
              GoTo 10
           ElseIf g_ier = -2 Then 'shutting down gracefully from oversize error
              GoTo 10
              End If
           End If
           
'
        Call DigitizerEnabling(1)
        
'        .Toolbar1.Buttons(36).Enabled = True 'enable digitizer '<<<<<<<<<<<<digi changes
'        .Toolbar1.Buttons(37).Enabled = True
'        .Toolbar1.Buttons(38).Enabled = True
'
'        .mnuDigitize.Enabled = True
'        .mnuDigiDeleteLastLine.Enabled = True
'        .mnuDigiDeleteLastPoint.Enabled = True
'        .mnuDigitizeContour.Enabled = True
'        .mnuDigitizeDeleteContour.Enabled = True
'        .mnuDigitizeDeleteLine.Enabled = True
'        .mnuDigitizeDeletPoint.Enabled = True
'        .mnuDigitizeEndContour.Enabled = True
'        .mnuDigitizeEndLine.Enabled = True
'        .mnuDigitizeEndPoint.Enabled = True
'        .mnuDigitizeLine.Enabled = True
'        .mnuDigitizePoint.Enabled = True
'        .mnuDigitizePointSameHeights.Enabled = True
        
        If buttonstate&(36) = 1 Then
          .Toolbar1.Buttons(36).value = tbrPressed
          If DigitizeMagvis Then 'reload this form to make it visible and with right dimensions
             Unload GDDigiMagfrm
             mnuDigitizer_Click
             End If
          End If
          
        If heights And Dir(dtmdir & "\dtm-map.loc") <> gsEmpty Then
           If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
              heights = True
              If Not JKHDTM Then InitializeDTM
              End If
           End If
           
        If UseNewDTM Then
           ier = OpenCloseBaseDTM(0)
           End If
        
     Else
     
                               
        .Toolbar1.Buttons(1).value = tbrUnpressed
         buttonstate&(2) = 0
        If topos = True Then .Toolbar1.Buttons(3).Enabled = True
        .Toolbar1.Buttons(8).Enabled = False 'disenable go-to buttons
        .Toolbar1.Buttons(10).Enabled = False 'disenable print maps
        .mnuPrintMap.Enabled = False
        .StatusBar1.Panels(1) = sEmpty
        .Toolbar1.Refresh

        'unload maps
        Call ShowGeoMap(1)
           
        If ImagePointFile Then 'close the point digitizing file
           Close #filnumImage%
           filnumImage% = 0
           ImagePointFile = False
           End If
        
'        .Toolbar1.Buttons(36).Enabled = False 'disenable the digitizer '<<<<<<<<<<<<digi changes
'        buttonstate&(36) = 0
'        .mnuDigitize.Enabled = False
        Call DigitizerEnabling(0)
'        If DigitizeMagvis Then
'          Unload GDDigiMagfrm
'          Unload GDform1
'          End If
        'close base dtm file
        ier = OpenCloseBaseDTM(1)
        
        End If
        
     End With

End Sub




Public Sub UnloadAllForms(Optional FormToIgnore _
  As String = sEmpty)

  Dim f As Form
  For Each f In Forms
    If f.Name <> FormToIgnore Then
      Unload f
      Set f = Nothing
    End If
  Next f
    
End Sub

Public Sub ReadWriteDefaultsandLink()

   Dim RepairInfo As Boolean
   On Error GoTo errdirhand
   
   GaussMethod = True 'Gaussian Elimination method is default method
   
   MinColorHeight = INIT_VALUE
   MaxColorHeight = -INIT_VALUE

   If Dir(direct$ + "\gdbinfo.sav") <> gsEmpty Then
   
        If infonum& > 0 Then
        Close #infonum&
        infonum& = 0
        End If

      infonum& = FreeFile
      Open direct$ + "\gdbinfo.sav" For Input As #infonum&
      Input #infonum&, doclin$
      Input #infonum&, dirNewDTM
      Input #infonum&, MinDigiEraserBrushSize
      Input #infonum&, NEDdir
      Input #infonum&, dtmdir
      Input #infonum&, ChainCodeMethod
      Input #infonum&, numDistContour, numDistLines, numSensitivity, numContours ' arcdir, mxddir
      Input #infonum&, PointCenterClick
      Input #infonum&, picnam$
      Input #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
      Input #infonum&, ReportPaths&, DigiSearchRegion, numMaxHighlight&, Save_xyz%
      Input #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
      Input #infonum&, IgnoreAutoRedrawError%
      Input #infonum&, UseNewDTM%, nOtherCheck%
      Input #infonum&, googledir, URL_OutCrop, URL_Well, kmldir, ASTERdir, DTMtype
      Input #infonum&, NX_CALDAT, NY_CALDAT
      Input #infonum&, RSMethod0, RSMethod1, RSMethod2
      Input #infonum&, ULPixX, ULPixY, LRPixX, LRPixY, LRGridX, LRGridY, ULGridX, ULGridY
      Input #infonum&, XStepITM, YStepITM, XStepDTM, YStepDTM, HalfAzi, StepAzi, Apprn, HeightPrecision, DigiConvertToMeters
      Close #infonum&
   
      If Trim$(picnam$) = sEmpty Then
         buttonstate&(2) = 0
         GDMDIform.Toolbar1.Buttons(2).Enabled = False
         GDMDIform.StatusBar1.Panels(1).Text = "No stored map file could be found, define one using the ""Options"" dialog (click first button on toolbar)..."
         End If
         
      If Dir(picnam$) = sEmpty Then
         'try adding app.path
         If Dir(App.Path & "\" & picnam$) <> sEmpty Then
            picnam$ = App.Path & "\" & picnam$
         Else
            buttonstate&(2) = 0
            GDMDIform.Toolbar1.Buttons(2).Enabled = False
            GDMDIform.StatusBar1.Panels(1).Text = "No stored map file could be found, define one using the ""Options"" dialog (click first button on toolbar)..."
            End If
         End If
         
         
    If MapUnits = 0 Or MapUnits = 1 Then
       MapUnits = 1#
       GDMDIform.Text3.ToolTipText = "Elevation (meters)"
       GDMDIform.Text7.ToolTipText = "Elevation (meters) at center of clicked point"
    ElseIf MapUnits = 0.30479999798832 Then
       GDMDIform.Text3.ToolTipText = "Elevation (feet)"
       GDMDIform.Text7.ToolTipText = "Elevation (feet) at center of clicked point"
    ElseIf MapUnits = 1.8288002 Then
       GDMDIform.Text3.ToolTipText = "Elevation (fathoms)"
       GDMDIform.Text7.ToolTipText = "Elevation (fathoms) at center of clicked point"
       End If

         
         
      If val(MinDigiEraserBrushSize) = 0 Then MinDigiEraserBrushSize = 1
      
      '*******set defaults for geologic map*******
      If Not Digitizing Then
         If lblX = sEmpty Then lblX = "ITMx"
         If LblY = sEmpty Then LblY = "ITMy"
         End If
         
      If RSMethod0 Then
        'determine coordinate conversion constants
        If LRPixX <> ULPixX Then
           PixToCoordX = (LRGeoX - ULGeoX) / (LRPixX - ULPixX)
'           RSMethod0 = False
        If ULPixY <> LRPixY Then
           PixToCoordY = (ULGeoY - LRGeoY) / (ULPixY - LRPixY)
'           RSMethod0 = False
           End If
        End If
      
'      If Installation_Type = 0 Then
'            If picnam$ = sEmpty Or IsNull(picnam$) Then
'               If Dir(direct$ + "\new5.bmp") <> sEmpty Then
'                  picnam$ = direct$ + "\new5.bmp"
'                  End If
'               End If
'            If IsNull(ulgeox) Then ulgeox = 80000#
'            If IsNull(ulgeoy) Then ulgeoy = 1300000#
'            If IsNull(lrgeox) Then lrgeox = 240000#
'            If IsNull(lrgeoy) Then lrgeoy = 880000#
'            If ulgeox = 0 And ulgeoy = 0 And lrgeox = 0 And lrgeoy = 0 Then
'               ulgeox = 80000#
'               ulgeoy = 1300000#
'               lrgeox = 240000#
'               lrgeoy = 880000#
'               End If
'            If IsNull(pixwi) Then pixwi = 1268
'            If IsNull(pixhi) Then pixhi = 3338
'            If pixwi = 0 Then pixwi = 1268
'            If pixhi = 0 Then pixhi = 3338
'            pixwi0 = pixwi
'            pixhi0 = pixhi
'
'            ULGeoX = ulgeox
'            ULGeoY = ulgeoy
'            LRGeoX = lrgeox
'            LRGeoY = lrgeoy
'
'      ElseIf Installation_Type = 1 Then
'         If picnam$ = sEmpty Or IsNull(picnam$) Then
'            If Dir(direct$ & "\IsraelShadedReliefMap.jpg") <> sEmpty Then
'               picnam$ = direct$ & "\IsraelShadedReliefMap.jpg"
'               End If
'            End If
            
        If IsNull(ULGeoX) Then ULGeoX = 0 '80000#
        If IsNull(ULGeoY) Then ULGeoY = 0 '1300000#
        If IsNull(LRGeoX) Then LRGeoX = 0 '240000#
        If IsNull(LRGeoY) Then LRGeoY = 0 '880000#
'        If ulgeox = 0 And ulgeoy = 0 And lrgeox = 0 And lrgeoy = 0 Then
'           ulgeox = 80000#
'           ulgeoy = 1300000#
'           lrgeox = 240000#
'           lrgeoy = 880000#
'           End If
        If IsNull(pixwi) Then pixwi = 1162
        If IsNull(pixhi) Then pixhi = 3046
        If pixwi = 0 Then pixwi = 1162
        If pixhi = 0 Then pixhi = 3046
        pixwi0 = pixwi
        pixhi0 = pixhi
        
rwdl100:
        End If
        
    If numcpt = 0 And LineElevColors& = 1 Then
         
         myfile = Dir(App.Path & "\rainbow.cpt")
         If myfile = sEmpty Then
            GoTo rwdl100
            End If
         
         '-----------------------load color palette--------------------------
         numpercent = -1
         numloop% = 0
         nowread = True
         num% = 0
         
         ReDim cpt(3, 0)
         
         filenum% = FreeFile
         Open App.Path & "\rainbow.cpt" For Input As #filenum%
         
         Do Until EOF(filenum%)
            Line Input #filenum%, doclin$
            colorattributes = Split(doclin$, " ")
            For i = 0 To 10
              cc$ = colorattributes(i)
              If Trim$(cc$) <> vbNullString Then
                 If numloop% = 0 Then
                    If val(cc$) >= numpercent Then
                        num% = val(cc$)
                        
                        If num% - 1 > UBound(cpt, 2) Then
                           ReDim Preserve cpt(3, UBound(cpt, 2) + 1)
                           End If
                        
                        cpt(0, num% - 1) = val(cc$)
                        numloop% = 1
                        numpercent = val(cc$)
                        nowread = True
                    Else
                        nowread = False
                        End If
                 ElseIf numloop% = 1 Then
                    If nowread Then cpt(1, num% - 1) = val(cc$)
                    numloop% = 2
                 ElseIf numloop% = 2 Then
                    If nowread Then cpt(2, num% - 1) = val(cc$)
                    numloop% = 3
                 ElseIf numloop% = 3 Then
                    If nowread Then cpt(3, num% - 1) = val(cc$)
                    numloop% = 0
                    nowread = False
                    Exit For
                    End If
                 End If
                 
                 numcpt = num%
                 
            Next i
         Loop
         Close #filenum%
         End If
        
      
     If RepairInfo Then 'repair the info (this provides back-compatibility
       'to old versions of the gdbinfo.sav files)
       If PointColor& = 0 Then PointColor& = 255 'red
       If LineColor& = 0 Then LineColor& = 65280 'green '65535 'yellow
       If ContourColor& = 0 Then ContourColor& = 10485760 'dark blue
       If RSColor& = 0 Then RSColor& = 65535 'yellow '65280 'green
      
'       If UnknownColor& = 0 Then UnknownColor& = 8388736 'purple
       If numMaxHighlight& = 0 Then numMaxHighlight& = 20000
       If SaveClose% = 0 Then SaveClose% = 1
       If Save_xyz% = 0 Then Save_xyz% = 1
'       If UseNewDTM% = 0 Then UseNewDTM% = 1
       If nOtherCheck% = 0 Then nOtherCheck% = 1
       If RSMethod1 = False And RSMethod2 = False And Not RSMethod0 Then
          RSMethod1 = True
          RSMethod2 = False
          RSMethod0 = False
          End If
'       If ASTERdir = "" Then ASTERdir = ""
'       If DTMtype = 0 Then DTMtype = 0
       IgnoreAutoRedrawError% = 0
       lblX = "lon." '"ITMx"
       LblY = "lat." '"ITMy"
       
       txtDistPixelSearch = DigiSearchRegion
       txtEraserBrushSize = MinDigiEraserBrushSize
       
'       If Dir(picnam$) <> sEmpty Then
'          'keep it, but check the other map settings
'          If ulgeox = 0 And ulgeoy = 0 And lrgeox = 0 And lrgeoy = 0 Then
'             ulgeox = 80000#
'             ulgeoy = 1300000#
'             lrgeox = 240000#
'             lrgeoy = 880000#
'             pixwi = 1268
'             pixhi = 3338
'             pixwi0 = pixwi
'             pixhi0 = pixhi
'             End If
'       ElseIf Dir(direct$ & "\new5.bmp") <> sEmpty Then
'          picnam$ = direct$ & "\new5.bmp"
'          ulgeox = 80000#
'          ulgeoy = 1300000#
'          lrgeox = 240000#
'          lrgeoy = 880000#
'          pixwi = 1268
'          pixhi = 3338
'          pixwi0 = pixwi
'          pixhi0 = pixhi
'       Else
'          picnam$ = sEmpty
'          ulgeox = 80000#
'          ulgeoy = 1300000#
'          lrgeox = 240000#
'          lrgeoy = 880000#
'          pixwi = 1268
'          pixhi = 3338
'          pixwi0 = pixwi
'          pixhi0 = pixhi
'          End If
       End If
    
     'now check these defaults
     '<<<<<<<<<<<<<<<comment out for this version>>>>>>>>>>>>>
     On Error GoTo dtmerror
     If Dir(NEDdir & "\Z000000.hgt") <> gsEmpty Then
        heights = True
        If DTMtype = 0 Then DTMtype = 2
     Else
        If DTMtype = 0 Then heights = False
        If NEDdir = sEmpty And Installation_Type = 1 Then 'try the default
           If Dir(direct$ & "\usa\Z000000.hgt") <> sEmpty Then
              NEDdir = direct$ & "\usa"
              If DTMtype = 0 Then DTMtype = 2
              heights = True
              End If
           End If
        End If
        
   If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
        If Dir(dtmdir & "\dtm-map.loc") <> gsEmpty Then
           heights = True
           If DTMtype = 0 Then DTMtype = 2
           If Not JKHDTM Then InitializeDTM
        Else
           If DTMtype = 0 Then heights = False
           If dtmdir = sEmpty And Installation_Type = 1 Then 'try default
              If Dir(direct$ & "\dtm\dtm-map.loc") <> sEmpty Then
                 dtmdir = direct$ & "\dtm"
                 If DTMtype = 0 Then DTMtype = 2
                 heights = True
                 End If
              End If
           End If
        End If
        
       If JKHDTM Then
          If val(XStepITM) = 0 Then XStepITM = 25
          If val(YStepITM) = 0 Then YStepITM = 30
       Else
          If val(XStepITM) = 0 Then XStepITM = 30
          If val(YStepITM) = 0 Then YStepITM = 30
          End If
          
       If val(XStepDTM) = 0 Then
          XStepDTM = 1#  '8.33333333333333E-04 / 3#
       Else
          XStepDTM = XStepDTM
          End If
       If val(YStepDTM) = 0 Then
          YStepDTM = 1#  '8.33333333333333E-04 / 3#
       Else
          YStepDTM = YStepDTM
          End If
       
        
     
'     On Error GoTo dtmerror
'     If Dir(dtmdir + "\dtm-map.loc") <> gsEmpty Then
'        heights = True
'        If DTMtype = 0 Then DTMtype = 2
'     Else
'        If DTMtype = 0 Then heights = False
'        If dtmdir = sEmpty And Installation_Type = 1 Then 'try the default
'           If Dir(direct$ & "\dtm\dtm-map.loc") <> sEmpty Then
'              dtmdir = direct$ & "\dtm"
'              If DTMtype = 0 Then DTMtype = 2
'              heights = True
'              End If
'           End If
'        End If
         
     If DTMtype = 0 Then
        If Dir(ASTERdir + "\N31E035.bil") <> gsEmpty Then
           heights = True
           If DTMtype = 0 Then DTMtype = 1
        Else
           If DTMtype = 0 Then heights = False
           If ASTERdir = sEmpty And Installation_Type = 1 Then 'try the default
              If Dir(direct$ & "\ASTER\N31E035.bil") <> sEmpty Then
                 ASTERdir = direct$ & "\ASTER"
                 If DTMtype = 0 Then DTMtype = 1
                 heights = True
                 End If
              End If
           End If
        End If
        
     If ((Mid$(LCase(lblX), 1, 3) <> "itm" And Mid$(LCase(LblY), 1, 3) <> "itm") And _
         (Mid$(LCase(lblX), 1, 3) <> "lon" And Mid$(LCase(LblY), 1, 3) <> "lat")) Then
         heights = False 'utm coord conversion not yet supported
         End If

        
chk1:
         
chk2:
         
         
chk3:
         
     'now default google earth directory
      defgoogle$ = Mid$(direct$, 1, 3) & "Program Files\Google\client\Google Earth"
      If Dir(defgoogle$ & "\googleearth.exe") <> sEmpty Then
         googledir = defgoogle$
         google = True
         URL_OutCrop = "http://maps.google.com/mapfiles/kml/pal4/icon49.png"
         URL_Well = "http://maps.google.com/mapfiles/kml/pal4/icon48.png"
         kmldir = direct$
         End If
         
   Else
   
   
      'now check these defaults
     
      '*******set defaults for geologic map*******
      If lblX = sEmpty Then lblX = "lon." '"ITMx"
      If LblY = sEmpty Then LblY = "lat." '"ITMy"
      
      If DigiSearchRegion = 0 Then DigiSearchRegion = 10
    
'      If Installation_Type = 0 Then
'            If picnam$ = sEmpty Or IsNull(picnam$) Then
'               If Dir(direct$ + "\new5.bmp") <> sEmpty Then
'                  picnam$ = direct$ + "\new5.bmp"
'                  End If
'               End If
'            If IsNull(ulgeox) Then ulgeox = 80000#
'            If IsNull(ulgeoy) Then ulgeoy = 1300000#
'            If IsNull(lrgeox) Then lrgeox = 240000#
'            If IsNull(lrgeoy) Then lrgeoy = 880000#
'            If ulgeox = 0 And ulgeoy = 0 And lrgeox = 0 And lrgeoy = 0 Then
'               ulgeox = 80000#
'               ulgeoy = 1300000#
'               lrgeox = 240000#
'               lrgeoy = 880000#
'               End If
'            If IsNull(pixwi) Then pixwi = 1268
'            If IsNull(pixhi) Then pixhi = 3338
'            If pixwi = 0 Then pixwi = 1268
'            If pixhi = 0 Then pixhi = 3338
'            pixwi0 = pixwi
'            pixhi0 = pixhi
'
'      ElseIf Installation_Type = 1 Then
'         If picnam$ = sEmpty Or IsNull(picnam$) Then
'            If Dir(direct$ & "\IsraelShadedReliefMap.jpg") <> sEmpty Then
'               picnam$ = direct$ & "\IsraelShadedReliefMap.jpg"
'               End If
'            End If
            
        If IsNull(ULGeoX) Then ULGeoX = 0 '80000#
        If IsNull(ULGeoY) Then ULGeoY = 0 '1300000#
        If IsNull(LRGeoX) Then LRGeoX = 0 '240000#
        If IsNull(LRGeoY) Then LRGeoY = 0 '880000#
'        If ulgeox = 0 And ulgeoy = 0 And lrgeox = 0 And lrgeoy = 0 Then
'           ulgeox = 80000#
'           ulgeoy = 1300000#
'           lrgeox = 240000#
'           lrgeoy = 880000#
'           End If
'        If IsNull(pixwi) Then pixwi = 1162
'        If IsNull(pixhi) Then pixhi = 3046
'        If pixwi = 0 Then pixwi = 1162
'        If pixhi = 0 Then pixhi = 3046
        pixwi0 = pixwi
        pixhi0 = pixhi
        
       If JKHDTM Then
          If val(XStepITM) = 0 Then XStepITM = 25
          If val(YStepITM) = 0 Then YStepITM = 30
       Else
          If val(XStepITM) = 0 Then XStepITM = 30
          If val(YStepITM) = 0 Then YStepITM = 30
          End If
          
       If val(XStepDTM) = 0 Then XStepDTM = 1#  '8.33333333333333E-04 / 3#
       If val(YStepDTM) = 0 Then YStepDTM = 1#  '8.33333333333333E-04 / 3#
       
        
        End If
     
     On Error GoTo dtmerror
     If Dir(NEDdir & "\Z000000.hgt") <> gsEmpty Then
        heights = True
        If DTMtype = 0 Then DTMtype = 2
     Else
        If DTMtype = 0 Then heights = False
        If NEDdir = sEmpty And Installation_Type = 1 Then 'try the default
           If Dir(direct$ & "\usa\Z000000.hgt") <> sEmpty Then
              NEDdir = direct$ & "\usa"
              If DTMtype = 0 Then DTMtype = 2
              heights = True
              End If
           End If
        End If
        
     If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
        If Dir(dtmdir + "\dtm-map.loc") <> gsEmpty Then
           heights = True
           If DTMtype = 0 Then DTMtype = 2
        Else
           If DTMtype = 0 Then heights = False
           If dtmdir = sEmpty And Installation_Type = 1 Then 'try the default
              If Dir(direct$ & "\dtm\dtm-map.loc") <> sEmpty Then
                 dtmdir = direct$ & "\dtm"
                 If DTMtype = 0 Then DTMtype = 2
                 heights = True
                 End If
              End If
           End If
         End If
         
     If DTMtype = 0 Then
        If Dir(ASTERdir + "\N31E035.bil") <> gsEmpty Then
           heights = True
           If DTMtype = 0 Then DTMtype = 1
        Else
           If DTMtype = 0 Then heights = False
           If ASTERdir = sEmpty And Installation_Type = 1 Then 'try the default
              If Dir(direct$ & "\ASTER\N31E035.bil") <> sEmpty Then
                 ASTERdir = direct$ & "\ASTER"
                 If DTMtype = 0 Then DTMtype = 1
                 heights = True
                 End If
              End If
           End If
        End If
           
     If ((Mid$(LCase(lblX), 1, 3) <> "itm" And Mid$(LCase(LblY), 1, 3) <> "itm") And _
         (Mid$(LCase(lblX), 1, 3) <> "lon" And Mid$(LCase(LblY), 1, 3) <> "lat")) Then
         heights = False 'utm coord conversion not yet supported
         End If
         
      On Error GoTo errdirhand
         
      If Dir(Mid$(direct$, 1, 3) & "Program Files\Google\client\Google Earth\googleearth.exe") <> sEmpty Then
         googledir = Mid$(direct$, 1, 3) & "Program Files\Google\client\Google Earth"
         google = True
         End If
         
      'default URL's for outcroppings and wells icons to use in the kml output file
      txtOutCropIcon = "http://maps.google.com/mapfiles/kml/pal4/icon49.png"
      URL_OutCrop = txtOutCropIcon
      txtWellIcon = "http://maps.google.com/mapfiles/kml/pal4/icon48.png"
      URL_Well = txtWellIcon
      
      If numMaxHighlight& = 0 Then numMaxHighlight& = 32767 'default number of allowed search results to plot set to maximum
      
'      If UseNewDTM% = 0 Then UseNewDTM% = 1 'replace null surface heights with dtm height
'      If nOtherCheck% = 0 Then nOtherCheck% = 1

 '///////////////////////////////////////////////////
      
      ''if default map is present, then generate info file
      'If Dir(direct$ + "\new5.bmp") <> gsEmpty Then
      '   infonum& = FreeFile
      '   Open direct$ + "\gdbinfo.sav" For Output As #infonum&
      '   Write #infonum&, "This file is used by the MapDigitizer program. Don't erase it!"
      '   Write #infonum&, sEmpty
      '   Write #infonum&, sEmpty
      '   Write #infonum&, sEmpty
      '   Write #infonum&, sEmpty
      '   Write #infonum&, sEmpty
      '   Write #infonum&, sEmpty, sEmpty
      '   Write #infonum&, sEmpty
      '   Write #infonum&, direct$ + "\new5.bmp"
      '   Write #infonum&, "ITMx", "ITMy", 80000, 240000, 1300000, 880000, 1268, 3338
      '   Write #infonum&, 0, 0, 20000, 1
      '   Write #infonum&, 255, 65535, 10485760, 65280, 8388736
      '   Close #infonum&
      '   End If
      
         
'      End If
            
     
     'generate map save file if not already present
'     If Dir(direct$ & "\gdbmap.sav") = sEmpty Then
'        infonum& = FreeFile
'        Open direct$ + "\gdbmap.sav" For Output As #infonum&
'        Write #infonum&, "This file is used by the MapDigitizer program. Don't erase it!"
'        If Dir(direct$ & "\IsraelShadedReliefMap.jpg") <> sEmpty Then
'           Write #infonum&, direct$ & "\IsraelShadedReliefMap.jpg", "ITMx", "ITMy", " 80000", " 240000", " 1300000", " 880000", " 1162", " 3046"
'           End If
'        If Dir(direct$ & "\IsraelSeismic-HazardMap.jpg") <> sEmpty Then
'           Write #infonum&, direct$ & "\IsraelSeismic-HazardMap.jpg", "ITMx", "ITMy", " 80000", " 240000", " 1305000", " 880000", " 1515", " 4048"
'           End If
'        If Dir(direct$ & "\IsraelMagneticsMap.jpg") <> sEmpty Then
'           Write #infonum&, direct$ & "\IsraelMagneticsMap.jpg", "ITMx", "ITMy", " 60000", " 240000", " 1300000", " 880000", " 1720", " 3970"
'           End If
'        If Dir(direct$ & "\IsraelGravityMap.jpg") <> sEmpty Then
'           Write #infonum&, direct$ & "\IsraelGravityMap.jpg", "ITMx", "ITMy", " 60000", " 240000", " 1300000", " 880000", " 1802", " 4096"
'           End If
'        If Dir(direct$ & "\Oilfields.jpg") <> sEmpty Then
'           Write #infonum&, direct$ & "\Oilfields.jpg", "ITMx", "ITMy", " 71000", " 240000", " 1320000", " 861000", " 484", " 1309"
'           End If
'        If Dir(direct$ & "\oil-wells-Israel.jpg") <> sEmpty Then
'           Write #infonum&, direct$ & "\oil-wells-Israel.jpg, "; ITMx; ", "; ITMy; ", "; 71000; ", "; 208225; ", "; 1169622; ", "; 1030940; ", "; 800; ", "; 810; ""
'           End If
'        If Dir(direct$ & "\Israel_Sinai.jpg") <> sEmpty Then
'           Write #infonum&, direct$ & "\Israel_Sinai.jpg, "; ITMx; ", "; ITMy; ", "; -146568; ", "; 253928; ", "; 1091269; ", "; 682642; ", "; 2319; ", "; 2456; ""
'           End If
'
'        PointColor& = 255
'        LineColor& = 65535
'        ContourColor& = 10485760
'        RSColor& = 65280
'        UnknownColor& = 8388736
'
'        Close #infonum&
'        End If
            
      '**********Link to Paleontological Database**********
'      'Link tables in paleontolgical data base on the server
'      'to dummy data base used by this program.
'      If dbdir2 <> gsEmpty Then LinkTables
'      '******************************************
'
'     '**********Link to Old Paleontological Database**********
'      'Link tables in old paleontolgical data base on the server
'      'to dummy data base used by this program.
'      If NEDdir <> gsEmpty Then
'         LinkTablesOld
'         LinkDBpiv
'         End If
         
'      If SearchDBs% = 0 Then
'         If linked And linkedOld Then
'            SearchDBs% = 1
'         ElseIf linked And Not linkedOld Then
'            SearchDBs% = 2
'         ElseIf Not linked And linkedOld Then
'            SearchDBs% = 3
'            End If
'         End If
         
      SaveClose% = 1 'always save new settings
'      Save_xyz% = 1 'save xyz data is default
            
      'now refresh recorded values
      
      If dtmdir = gsEmpty Then
         If Dir(direct$ & "\dtm\dtm-map.loc") <> gsEmpty Then
            dtmdir = direct$ & "\dtm"
            End If
         End If
      
        If infonum& > 0 Then
           Close #infonum&
           infonum& = 0
           End If
      
      infonum& = FreeFile
      Open direct$ + "\gdbinfo.sav" For Output As #infonum&
      Write #infonum&, "This file is used by the MapDigitizer program. Don't erase it!"
      Write #infonum&, dirNewDTM
      Write #infonum&, MinDigiEraserBrushSize
      Write #infonum&, NEDdir
      Write #infonum&, dtmdir
      Write #infonum&, ChainCodeMethod
      Write #infonum&, numDistContour, numDistLines, numSensitivity, numContours ' arcdir, mxddir
      Write #infonum&, PointCenterClick
      Write #infonum&, picnam$
      Write #infonum&, lblX, LblY, ULGeoX, LRGeoX, ULGeoY, LRGeoY, pixwi, pixhi, MapUnits
      Write #infonum&, ReportPaths&, DigiSearchRegion, numMaxHighlight&, Save_xyz%
      Write #infonum&, PointColor&, LineColor&, ContourColor&, RSColor&, LineElevColors&
      Write #infonum&, IgnoreAutoRedrawError%
      Write #infonum&, UseNewDTM%, nOtherCheck%
      Write #infonum&, googledir, URL_OutCrop, URL_Well, kmldir, ASTERdir, DTMtype
      Write #infonum&, NX_CALDAT, NY_CALDAT
      Write #infonum&, RSMethod0, RSMethod1, RSMethod2
      Write #infonum&, ULPixX, ULPixY, LRPixX, LRPixY, LRGridX, LRGridY, ULGridX, ULGridY
      Write #infonum&, XStepITM, YStepITM, XStepDTM, YStepDTM, HalfAzi, StepAzi, Apprn, HeightPrecision, DigiConvertToMeters
      Close #infonum&
    
     'Geologic map parameters
      picnam0$ = picnam$
      x10 = ULGeoX
      y10 = ULGeoY
      x20 = LRGeoX
      y20 = LRGeoY
      pixwi0 = pixwi
      pixhi0 = pixhi
       
      Close #infonum&
      
      'if directory name is only letter, like d:\ then truncate
      If InStr(dirNewDTM, "\") <> 0 And Len(dirNewDTM) = 3 Then dirNewDTM = Mid$(dirNewDTM, 1, 2)
'      If InStr(dbdir2, "\") <> 0 And Len(dbdir2) = 3 Then dbdir2 = Mid$(dbdir2, 1, 2)
      If InStr(dtmdir, "\") <> 0 And Len(dtmdir) = 3 Then dtmdir = Mid$(dtmdir, 1, 2)
      If InStr(NEDdir, "\") <> 0 And Len(NEDdir) = 3 Then NEDdir = Mid$(NEDdir, 1, 2)
      If InStr(ASTERdir, "\") <> 0 And Len(ASTERdir) = 3 Then ASTERdir = Mid$(ASTERdir, 1, 2)
'      If InStr(topodir, "\") <> 0 And Len(topodir) = 3 Then topodir = Mid$(topodir, 1, 2)
'      If InStr(arcdir, "\") <> 0 And Len(arcdir) = 3 Then arcdir = Mid$(arcdir, 1, 2)
'      If InStr(accdir, "\") <> 0 And Len(accdir) = 3 Then accdir = Mid$(accdir, 1, 2)
      If InStr(googledir, "\") <> 0 And Len(googledir) = 3 Then googledir = Mid$(googledir, 1, 2)
      If InStr(kmldir, "\") <> 0 And Len(kmldir) = 3 Then kmldir = Mid$(kmldir, 1, 2)
      
    If MapUnits = 0 Or MapUnits = 1 Then
       MapUnits = 1#
       GDMDIform.Text3.ToolTipText = "Elevation (meters)"
       GDMDIform.Text7.ToolTipText = "Elevation (meters) at center of clicked point"
    ElseIf MapUnits = 0.30479999798832 Then
       GDMDIform.Text3.ToolTipText = "Elevation (feet)"
       GDMDIform.Text7.ToolTipText = "Elevation (feet) at center of clicked point"
    ElseIf MapUnits = 1.8288002 Then
       GDMDIform.Text3.ToolTipText = "Elevation (fathoms)"
       GDMDIform.Text7.ToolTipText = "Elevation (fathoms) at center of clicked point"
       End If
      
chk4:
'   On Error GoTo errdirhand '<<<<<<<<<<<<<<<<comment out for this version>>>>>>>>>>>>>>>>
'   If heights = True And DTMtype = 2 Then
'      'load in parameters for DTM heights
'      InitializeDTM
'      End If
      
chk5: On Error GoTo googleerror
      If Dir(googledir + "\googleearth.exe") <> gsEmpty Then
         google = True
      Else
         google = False
         End If
         
      If Trim$(kmldir) = sEmpty Then kmldir = direct$
      
      
Exit Sub
            
      
'***********ERROR CHECKING ROUTINES***************
dtmerror:
   heights = False 'the path to the dtm is in error
   Err.Clear
   GoTo chk1
   
topoerror:
   topos = False 'the path to the topo maps is in error
   Err.Clear
   GoTo chk2
   
accerror:
   If Err.Number = 5 Then 'password hasn't been registered yet
      Resume Next
      End If
   acc = False 'the path to MSAccess.exe is in error
   Err.Clear
   GoTo chk3
   
arcerror:
   arcs = False 'the path to ArcMap.exe is in error
   Err.Clear
   GoTo chk4
   
googleerror:
   google = False 'the path to Google Earth is in error
   Err.Clear
'   GoTo main150
   
Exit Sub

errdirhand:
       'something bad is wrong with something else
       'so show error message and ignore all defined paths
        
        If Err.Number = 75 And errpal& = 1 Then
            Screen.MousePointer = vbDefault
            GDsplash.Visible = False
            Unload GDsplash
            SplashVis = False
            errpal& = 0
            'can't kill the old temporary direct$ & "\pal_dt_tmp.mdb directory
            MsgBox "Can't erase the old temporary database: " & vbLf & _
                   direct$ & "\pal_dt_tmp.mdb!" & vbLf & _
                   "Exit the program, erase that file, and start the program again.", _
                   vbOKOnly + vbExclamation, "MapDigitizer"
        ElseIf Err.Number = 62 Then 'old gbinfo.sav file is corrupted--repair it
           'repair as much as possible
           RepairInfo = True
           Resume Next
        Else 'something unexpected
            Screen.MousePointer = vbDefault
            GDsplash.Visible = False
            Unload GDsplash
            SplashVis = False
            Close
            MsgBox "Encountered unexpected error #: " & Err.Number & vbLf & _
                Err.Description & vbLf & vbLf & _
                "Warning: Loading of paths and options was not completed", _
                vbCritical + vbOKOnly, "MapDigitizer"
'            linked = False
'            heights = False
'            topos = False
'            arcs = False
'            acc = False
            Err.Clear
            GoTo chk5
       
        End If

End Sub
Sub DigitizerEnabling(mode%)

If mode% = 1 Then
   With GDMDIform
        If RSMethod0 And ULGeoX <> LRGeoX And ULGeoY <> LRGeoY _
           And ULPixX <> LRPixX And ULPixY <> LRPixY Then .Toolbar1.Buttons(8).Enabled = True 'goto
           
        XminC = 0
        YminC = 0
        XmaxC = 0
        YmaxC = 0
        
        .combContour.Visible = False
           
        .Toolbar1.Buttons(10).Enabled = True
        .Toolbar1.Buttons(15).Enabled = True
        .Toolbar1.Buttons(28).Enabled = False 'report
        .Toolbar1.Buttons(30).Enabled = False 'save
        .Toolbar1.Buttons(3).Enabled = True 'open xyz file (formally the MapTopo button)
        
        If Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" And Trim$(googledir) <> sEmpty Then .Toolbar1.Buttons(33).Enabled = True 'Google
        .Toolbar1.Buttons(36).Enabled = True '
        .Toolbar1.Buttons(37).Enabled = True 'enable digitizing
        .Toolbar1.Buttons(38).Enabled = True 'extend grids
        .Toolbar1.Buttons(39).Enabled = True 'rubber sheeting
        .Toolbar1.Buttons(40).Enabled = False 'eraser
        .Toolbar1.Buttons(41).Enabled = False 'sweep off digitized points
        .Toolbar1.Buttons(42).Enabled = False 'edit digitizied points
'        .Toolbar1.Buttons(43).Enabled = True 'Hardy
        .Toolbar1.Buttons(44).Enabled = True 'openXYZ
        
'        .Toolbar1.Buttons(45).Enabled = True 'merge files to create DTM
         If Installation_Type = 1 Then
            .Toolbar1.Buttons(47).Visible = True
            .Toolbar1.Buttons(47).Enabled = True 'GTCO tablet interface
            End If
        
'        .mnuDigitize.Enabled = True
'        .mnuDigitizeRubberSheeting.Enabled = True
'        .mnuDigiDeleteLastLine.Enabled = True
'        .mnuDigiDeleteLastPoint.Enabled = True
'        .mnuDigitizeContour.Enabled = True
''        .mnuDigitizeDeleteContour.Enabled = True
'        .mnuDigitizeDeleteLine.Enabled = True
'        .mnuDigitizeDeletPoint.Enabled = True
'        .mnuDigitizeEndContour.Enabled = True
'        .mnuDigitizeEndLine.Enabled = True
'        .mnuDigitizeEndPoint.Enabled = True
'        .mnuDigitizeLine.Enabled = True
'        .mnuDigitizePoint.Enabled = True
'        .mnuDigitizePointSameHeights.Enabled = True
'        .mnuDigiExtendGrid.Enabled = True
        
        .mnuDigitize.Enabled = True
'        .mnuEraser.Enabled = True
'        .mnuDigitizeHardy.Enabled = True
'        .mnuDigiSweep.Enabled = True
        
        'initialize digitizer mouse coordinates
        digiextendgrid_last.x = INIT_VALUE
        digiextendgrid_last.Y = INIT_VALUE
        digiextendgrid_begin.x = INIT_VALUE
        digiextendgrid_begin.Y = INIT_VALUE
        
'        If LRPixX = 0 Then LRPixX = pixwi
'        If LRPixY = 0 Then LRPixY = pixhi
        
    End With
    
'    'make memory copy of geo map file to be used for the eraser tool
'    Set oGestionImageSrc.PictureBox = GDform1.Picture2
    
ElseIf mode% = 0 Then
    With GDMDIform
        .Toolbar1.Buttons(3).Enabled = False
        .Toolbar1.Buttons(8).Enabled = False
        .Toolbar1.Buttons(10).Enabled = False
        .Toolbar1.Buttons(15).Enabled = False
        .Toolbar1.Buttons(28).Enabled = False
        .Toolbar1.Buttons(30).Enabled = False
        .Toolbar1.Buttons(33).Enabled = False
        .Toolbar1.Buttons(34).Enabled = False
        .Toolbar1.Buttons(36).Enabled = False 'disenable digitizer '<<<<<<<<<<<<digi changes
        .Toolbar1.Buttons(37).Enabled = False
        .Toolbar1.Buttons(38).Enabled = False
        .Toolbar1.Buttons(39).Enabled = False
        .Toolbar1.Buttons(40).Enabled = False
        .Toolbar1.Buttons(41).Enabled = False
        .Toolbar1.Buttons(42).Enabled = False
        .Toolbar1.Buttons(43).Enabled = False
        .Toolbar1.Buttons(44).Enabled = False
        .Toolbar1.Buttons(45).Enabled = False
        .Toolbar1.Buttons(49).Enabled = False
        .Toolbar1.Buttons(50).Enabled = False
        .Toolbar1.Buttons(51).Enabled = False
        If Installation_Type = 1 Then .Toolbar1.Buttons(47).Enabled = False
        
        XminC = 0
        YminC = 0
        XmaxC = 0
        YmaxC = 0
        
         buttonstate&(3) = 0
         buttonstate&(15) = 0
         buttonstate&(28) = 0
         buttonstate&(30) = 0
         buttonstate&(34) = 0
         buttonstate&(36) = 0
         buttonstate&(37) = 0
         buttonstate&(38) = 0
         buttonstate&(39) = 0
         buttonstate&(40) = 0
         buttonstate&(41) = 0
         buttonstate&(42) = 0
         buttonstate&(43) = 0
         buttonstate&(44) = 0
         buttonstate&(45) = 0
         buttonstate&(46) = 0
         buttonstate&(50) = 0
         buttonstate&(51) = 0
         buttonstate&(52) = 0
         
         If Installation_Type = 1 Then buttonstate&(47) = 0
         
         .Toolbar1.Buttons(3).value = tbrUnpressed
         .Toolbar1.Buttons(15).value = tbrUnpressed
         .Toolbar1.Buttons(28).value = tbrUnpressed
         .Toolbar1.Buttons(30).value = tbrUnpressed
         .Toolbar1.Buttons(36).value = tbrUnpressed
         .Toolbar1.Buttons(37).value = tbrUnpressed
         .Toolbar1.Buttons(38).value = tbrUnpressed
         .Toolbar1.Buttons(39).value = tbrUnpressed
         .Toolbar1.Buttons(40).value = tbrUnpressed
         .Toolbar1.Buttons(41).value = tbrUnpressed
         .Toolbar1.Buttons(42).value = tbrUnpressed
         .Toolbar1.Buttons(43).value = tbrUnpressed
         .Toolbar1.Buttons(44).value = tbrUnpressed
         .Toolbar1.Buttons(45).value = tbrUnpressed
         .Toolbar1.Buttons(46).value = tbrUnpressed
         .Toolbar1.Buttons(52).value = tbrUnpressed
         If Installation_Type = 1 Then .Toolbar1.Buttons(47).Enabled = False
        
        .combContour.Visible = False
        
        .mnuDigitize.Enabled = False
        .mnuDigitizeRubberSheeting.Enabled = False
        .mnuDigiDeleteLastLine.Enabled = False
        .mnuDigiDeleteLastPoint.Enabled = False
        .mnuDigitizeContour.Enabled = False
'        .mnuDigitizeDeleteContour.Enabled = False
        .mnuDigitizeDeleteLine.Enabled = False
        .mnuDigitizeDeletPoint.Enabled = False
        .mnuDigitizeEndContour.Enabled = False
        .mnuDigitizeEndLine.Enabled = False
        .mnuDigitizeEndPoint.Enabled = False
        .mnuDigitizeLine.Enabled = False
        .mnuDigitizePoint.Enabled = False
        .mnuDigitizePointSameHeights.Enabled = False
        .mnuDigiExtendGrid.Enabled = False
        .mnuEraser.Enabled = False
        .mnuDigitizeHardy.Enabled = False
        .mnuDigiSweep.Enabled = False
        
        DigiRS = False
        DigitizeOn = False
        DigitizeContour = False
        DigitizePoint = False
        DigitizeLine = False
        DigiRubberSheeting = False
        DigitizeExtendGrid = False
        DigiExtendFirstPoint = False
        DigitizerSweep = False
        DigiTableWorksOpen = False
        DigitizeHardy = False
        DigiReDrawContours = False
        Belgier_Smoothing = False
        DigiReDrawContours = False
        GenerateContours = False
        DTMcreating = False
        BasisDTMheights = False
        
        'close any open files
        Close
        
        GDMDIform.Label1 = "XPix"
        GDMDIform.Label5 = "XPix"
        GDMDIform.Label2 = "YPix"
        GDMDIform.Label6 = "YPix"
        
        GDMDIform.Text3.Visible = False
        GDMDIform.Label3.Visible = False
        GDMDIform.Text7.Visible = False
        GDMDIform.Label7.Visible = False
            
        GDMDIform.Text4.Visible = False
        GDMDIform.Label4.Visible = False
        
        If Not DigiGDIfailed Then
           ReDim m_byImage(0) 'reclaim eraser buffer memory
        Else
           If DigiPicFileOpened And Picfilnum% > 0 Then Close #Picfilnum% 'close color file
           End If
        
        If DigitizeMagvis Then
          Unload GDDigiMagfrm
          Unload GDform1
          End If
          
        If GDRSfrmVis Then
           Unload GDRSfrm
           End If
   
        If RSopenedfile Then
           Close #RSfilnum%
           numRS = 0
           ReDim RS(numRS) 'reclaim memory
           End If
           
        InitDigiGraph = False
        MinColorHeight = INIT_VALUE
        MaxColorHeight = -INIT_VALUE
        
    End With
    End If
End Sub
Sub DigiMag() 'opens and closes magnification window
    If buttonstate&(36) = 0 Then
       buttonstate&(36) = 1
       
       If Not DigitizeMagvis Then 'magnify screen
          DigitizeMagInit = True
          GDDigiMagfrm.Visible = True
          End If
       
    Else
       buttonstate&(36) = 0
       GDMDIform.Toolbar1.Buttons(36).value = tbrUnpressed
       Unload GDDigiMagfrm
'       If DigitizePadVis Then Unload GDDigitizerfrm
       
      'renable blinking
       GDMDIform.CenterPointTimer.Enabled = True
       ce& = 1
      
       End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : OpenXYZfile
' Author    : Dr-John-K-Hall
' Date      : 2/22/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub OpenXYZfile() 'opens an xyz file and plots the contours
    
    Dim x() As Double, Y() As Double, Z() As Double
    Dim Xcoord() As Double, Ycoord() As Double
    Dim ht() As Double, htf() As Single, hts() As Integer
    Dim np As Long
    Dim xmin As Double, xmax As Double
    Dim ymin As Double, ymax As Double
    Dim zmin As Double, zmax As Double
    Dim i As Long, j As Long
    Dim XX As Double, YY As Double, zz As Double
    Dim nc As Integer
    Dim contour() As Double
    Dim ncols As Long, nrows As Long
    
    Dim kmxo As Double
    kmxo = -INIT_VALUE

    Dim ContourInterval As Integer
    ContourInterval = val(GDMDIform.combContour.Text) '5 '2 '10 '5 '10 '100 'contour intervals in height units
    
   On Error GoTo OpenXYZfile_Error

'    If buttonstate&(42) = 0 And GeoMap Or TopoMap Then
'       buttonstate&(42) = 1
       
        CommonDialog1.CancelError = True
        CommonDialog1.Filter = "xyz files (*.xyz)|*.xyz|all files (*.*)|*.*"
        CommonDialog1.FileName = App.Path & "\*.xyz"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.ShowOpen
        inputfile$ = CommonDialog1.FileName
        
        If Dir(inputfile$) <> sEmpty Then
        
           'read file to determine the number of cols
           GDMDIform.StatusBar1.Panels(1).Text = "Determining size of xyz file, please wait...."
           filnum% = FreeFile
           Open inputfile$ For Input As #filnum%
           Screen.MousePointer = vbHourglass
           
           Do Until EOF(filnum%)
              Line Input #filnum%, doclin$
              np = np + 1
'              DoEvents
           Loop
           Close #filnum%
           Screen.MousePointer = vbDefault
           
           'read file, stuff data into arrays, and plot the contours
           xmin = INIT_VALUE
           xmax = -INIT_VALUE
           ymin = INIT_VALUE
           ymax = -INIT_VALUE
           zmin = INIT_VALUE
           zmax = -INIT_VALUE
           
           GDMDIform.StatusBar1.Panels(1).Text = "Loading up xyz file, please wait...."
           Screen.MousePointer = vbHourglass
           
            '------------------progress bar initialization
            With GDMDIform
                 '------fancy progress bar settings---------
                 .picProgBar.AutoRedraw = True
                 .picProgBar.BackColor = &H8000000B 'light grey
                 .picProgBar.DrawMode = 10
               
                 .picProgBar.FillStyle = 0
                 .picProgBar.ForeColor = &H400000 'dark blue
                 .picProgBar.Visible = True
            End With
            pbScaleWidth = 100
            '-------------------------------------------------
            
            Call UpdateStatus(GDMDIform, 1, 0)
            
           Dim npts As Long
           npnts = 0
      
           kmxo = -INIT_VALUE
           
           filnum% = FreeFile
           Open inputfile$ For Input As #filnum%
           Do Until EOF(filnum%)
              Input #filnum%, XX, YY, zz
              xmin = min(xmin, XX)
              xmax = Max(xmax, XX)
              ymin = min(ymin, YY)
              ymax = Max(ymax, YY)
              zmin = min(zmin, zz)
              zmax = Max(zmax, zz)
              
              nrows = nrows + 1
              npnts = npnts + 1
              
              If kmxo <> XX Then
                 ncols = ncols + 1
                 kmxo = XX
                 nrows = 0
                 Call UpdateStatus(GDMDIform, 1, CLng(npnts * 100 / np))
                 End If
                 
              DoEvents
           Loop
           nrows = nrows + 1
           Close #filnum%
          
           GDMDIform.picProgBar.Visible = False
           
           If xmax > pixwi Or ymax > pixhi Then
              Call MsgBox("This file is not a topo_pixel.xyz type file or" _
                          & vbCrLf & "its pixel dimensions don't fit the current map!" _
                          , vbExclamation, "Plot error")
              Exit Sub
              End If
                      
           
           ReDim Xcoord(ncols - 1)
           ReDim Ycoord(nrows - 1)
           
           If HeightPrecision = 0 Then
                ReDim hts(ncols - 1, nrows - 1)
                ReDim htf(0)
                ReDim ht(0)
                
                filnum% = FreeFile
                Open inputfile$ For Input As #filnum%
                Call UpdateStatus(GDMDIform, 1, 0)
                For i = 0 To ncols - 1
                    For j = 0 To nrows - 1
                       Input #filnum%, Xcoord(i), Ycoord(j), hts(i, j)
                    Next j
                    If ncols > 0 Then Call UpdateStatus(GDMDIform, 1, CLng((i + 1) * 100 / ncols))
                Next i
                Close #filnum%
           ElseIf HeightPrecision = 1 Then
                ReDim htf(ncols - 1, nrows - 1)
                ReDim ht(0)
                ReDim hts(0)
                
                filnum% = FreeFile
                Open inputfile$ For Input As #filnum%
                Call UpdateStatus(GDMDIform, 1, 0)
                For i = 0 To ncols - 1
                    For j = 0 To nrows - 1
                       Input #filnum%, Xcoord(i), Ycoord(j), htf(i, j)
                    Next j
                    If np > 0 Then Call UpdateStatus(GDMDIform, 1, (i + 1) / ncols)
                Next i
                Close #filnum%
           ElseIf HeightPrecision = 2 Then
                ReDim ht(ncols - 1, nrows - 1)
                ReDim htf(0)
                ReDim hts(0)
                
                filnum% = FreeFile
                Open inputfile$ For Input As #filnum%
                Call UpdateStatus(GDMDIform, 1, 0)
                For i = 0 To ncols - 1
                    For j = 0 To nrows - 1
                       Input #filnum%, Xcoord(i), Ycoord(j), ht(i, j)
                    Next j
                    If np > 0 Then Call UpdateStatus(GDMDIform, 1, (i + 1) / np)
                Next i
                Close #filnum%
                End If
                
'            'now invert the y coordinates
'            Call UpdateStatus(GDMDIform, 1, 0)
'            Dim tmpY() As Double
'            ReDim tmpY(nrows - 1)
'            For j = 0 To nrows - 1
'                tmpY(nrows - j - 1) = Ycoord(j)
'                Call UpdateStatus(GDMDIform, 1, CLng(100 * j / (nrows - 1)))
'            Next j
'            Call UpdateStatus(GDMDIform, 1, 0)
'            For j = 0 To nrows - 1
'               Ycoord(j) = tmpY(j)
'               Call UpdateStatus(GDMDIform, 1, CLng(100 * j / (nrows - 1)))
'            Next j
'            ReDim tmpY(0) 'reclaim memory
    
            '-------------------generate contours----------------------------
            GDMDIform.StatusBar1.Panels(1).Text = "Generating and plotting contours, please wait......"
            numContourPoints = 0 'zero contour lines array
            ReDim ContourPoints(numContourPoints)
            ReDim contour(0) 'zero contour color array
            
            GDMDIform.combContour.Visible = True
            If numContours > 0 Then
               GDMDIform.combContour.Text = str(numContours)
            Else
               GDMDIform.combContour.ListIndex = 6 '(10 meters as default) '2 '(3 meters as default)
               End If
            
            ContourInterval = val(GDMDIform.combContour.Text) '2 '10 '5 '10 '100 'contour intervals in height units
            
            nc = Int((zmax - zmin) / ContourInterval)
        
            For i = 1 To nc
               If i > 0 Then
                  ReDim Preserve contour(i)
                  End If
               contour(i - 1) = zmin + (i - 1) * ContourInterval
            Next i
            
            ier = ReDrawMap(0)
            
            ier = conrec(GDform1.Picture2, ht, htf, hts, Xcoord, Ycoord, nc, contour, 0, ncols - 1, 0, nrows - 1, xmin, ymin, xmax, ymax, 0)
            If ier = -1 Then
               Screen.MousePointer = vbDefault
               Call MsgBox("Palette file: rainbow.cpt is missing in the program directory." _
                          & vbCrLf & "" _
                          & vbCrLf & "Contours won't be drawn" _
                          , vbExclamation, "Hardy contours")
               End If
               
            Screen.MousePointer = vbDefault
            
            End If
        
'    Else
'       buttonstate&(42) = 0
'       GDMDIform.Toolbar1.Buttons(42).value = tbrUnpressed
'       End If

   On Error GoTo 0
   Exit Sub

OpenXYZfile_Error:

    If filnum% > 0 Then Close #filnum%
    If filtopo% > 0 Then Close #filtopo%
    Screen.MousePointer = vbDefault
    GDMDIform.picProgBar.Visible = False

End Sub

Private Sub mnuOpen_Click()

'This routine replaces the VB CommonDialog control with a
'MultiSelect GetOpenFileName Common Dialog API
'(much of this code is from "www.mvps.org/vbnet/code/")

   Dim sFilters As String
   
   Dim pos As Long
   Dim buff As String
   Dim sLongname As String
   Dim sShortname As String
   Dim sRealPath As String
   
   'string of filters for the dialog box
   sFilters = "Csv documents (*.csv) " & vbNullChar & "*.csv" & vbNullChar & _
              "Text documents (*.txt)" & vbNullChar & "*.txt" & vbNullChar & _
              "Rel files (*.rel)" & vbNullChar & "*.rel" & vbNullChar & _
              "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
              
   With OFN
      'size of the OFN structure
      .nStructSize = Len(OFN)
      
       'window owning the dialog
      .hWndOwner = frmSetCond.hwnd
      
      'filters (patterns) for the dropdown combo
      .sFilter = sFilters
      
      'index to the default filter
      .nFilterIndex = 4
      
      'default filename, plus additional padding
      'for the user's final selection(s).  Must be
      'double-null terminated
      .sFile = "test.csv" & Space$(2048) & vbNullChar & vbNullChar
      
      'the size of the buffer
      .nMaxFile = Len(.sFile)
      
      'default extension applied to
      'file if it has no extension
      .sDefFileExt = "bas" & vbNullChar & vbNullChar
      
      'space fot he file title if a single selection
      'made, double-null terminated, and its size
      .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
      
      'starting folder, double-null terminated
      .nMaxTitle = Len(OFN.sFileTitle)
      
      'the dialog title
      If directPlot$ = "" Then directPlot$ = CurDir
      .sInitialDir = directPlot$ & vbNullChar & vbNullChar
      
      'default open flags and multiselect
      .sDialogTitle = "(Multi)select file(s) for plotting"
      .Flags = OFS_FILE_OPEN_FLAGS Or _
             OFN_ALLOWMULTISELECT
   End With
   
   If GetOpenFileName(OFN) Then
      'remove trailing pair of termnating nulls
      'and trim returned file string
      buff = Trim$(left$(OFN.sFile, Len(OFN.sFile) - 2))
      
      'Show the members of the returned sFile string
      'It path is larger than 3 characters, only show
      'most inner path (but record the rest in the file buffer)
      Dim sTemp As String, sPath As String, sShortPath As String
      Dim i%, numlist%, pos1%, MultiSelectPath As Boolean, sPath0 As String
      Dim sDriveLetter As String, MaxDirLen As Integer
      numlist% = 0
      MaxDirLen = Int(flxlstFiles.Width / 70) - 30
      
      sPath0 = left$(OFN.sFile, OFN.nFileOffset)
      
      If InStr(sPath0, vbNullChar) <> 0 Then
         'this is multiselect path without final "\"
         'first trim off vbnullchar
         sPath = TrimNull(sPath0)
         If Mid$(sPath, Len(sPath), 1) <> "\" Then sPath = sPath & "\"
         MultiSelectPath = True
      Else
         'this is single path with final "\"
         sPath = sPath0
         End If
         
      'record this plot information
      'if has final "\" then remove it
      directPlot$ = sPath
      If Mid$(directPlot$, Len(directPlot$), 1) = "\" Then
         directPlot$ = Mid$(directPlot$, 1, Len(directPlot$) - 1)
         End If
      Dim filplt%
      filplt% = FreeFile
      Open App.Path & "\PlotDirec.txt" For Output As #filplt%
      Write #filplt%, "This file is used by Plot. Don't erase it!"
      Write #filplt%, direct$, directPlot$, dirWordpad
      Close #filplt%
      
      
      'determine short form of this path consisting of
      'the innermost directory
      Call ShortPath(sPath, MaxDirLen, sShortPath, sRealPath)
      
      Do While Len(buff) > 3
         sTemp = StripDelimitedItem(buff, vbNullChar)
         If MultiSelectPath Then
            If Mid$(sTemp, 1, Len(sPath0)) <> sPath0 Then
                numfiles% = numfiles% + 1
                
                'here is where to add combo showing files
                'redimension plotinfo array
                ReDim Preserve PlotInfo(7, numfiles%)
'
'                'list everything but the directory path
'                'lstFiles.AddItem sShortPath & sTemp
'                flxlstFiles.AddItem sShortPath & sTemp
'                flxlstFiles.Refresh
                
                ReDim Preserve Files(numfiles%)
                Files(numfiles% - 1) = sPath & "\" & TrimNull(sTemp)
                End If
         Else
            'don't repeat the directory
             numfiles% = numfiles% + 1
             
             'show files
             'redimension plotinfo array
             ReDim Preserve PlotInfo(7, numfiles%)
'
'             'lstFiles.AddItem sShortPath & Mid$(sTemp, Len(sPath) + 2, Len(sTemp) - Len(sPath) - 1)
'             flxlstFiles.AddItem sShortPath & Mid$(sTemp, Len(sPath) + 2, Len(sTemp) - Len(sPath) - 1)
'             flxlstFiles.Refresh
             ReDim Preserve Files(numfiles%)
             Files(numfiles% - 1) = TrimNull(sTemp)
             End If
      Loop
   End If
     
   If numfiles% > 0 Then
      mnuSave.Enabled = True
      cmdShowEdit.Enabled = True
      cmdWizard.Enabled = True
      cmdAll.Enabled = True
      cmdClear.Enabled = True
      
        chkOrigin.Enabled = True
        chkGridLine.Enabled = True
        txtX0.Enabled = True
        txtX1.Enabled = True
        txtValueY0.Enabled = True
        txtValueY1.Enabled = True
        txtXTitle.Enabled = True
        txtYTitle.Enabled = True
        txtValueX0.Enabled = True
        txtValueX1.Enabled = True
        fraLayout.Enabled = True
        frmSetCond.lblEndX1.Enabled = True
        frmSetCond.lblEndY.Enabled = True
        frmSetCond.lblGridLine.Enabled = True
        frmSetCond.lblIndexEnd.Enabled = True
        frmSetCond.lblOrigin.Enabled = True
        frmSetCond.lblIndexStart.Enabled = True
        frmSetCond.lblStartX0.Enabled = True
        frmSetCond.lblStartY.Enabled = True
        frmSetCond.lblXTitle.Enabled = True
        frmSetCond.lblYTitle.Enabled = True
      
      End If
      
End Sub

Public Sub mnuSave_Click()

    If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(14, 1) 'show the right mode in the digitizer form

    SaveExcel 'save the search results
End Sub
Public Sub mnuTablet()

    If Installation_Type = 0 Then Exit Sub
    
    If (TopoMap Or GeoMap) Then
    
          If buttonstate&(47) = 0 Then
             buttonstate&(47) = 1
             GDMDIform.Toolbar1.Buttons(47).value = tbrnpressed
             Load TabConSample_VB_Form
             
          If TabletControlVis Then
    
              If buttonstate&(37) = 0 Then 'activate digitize button
                 buttonstate&(37) = 1
                 GDMDIform.Toolbar1.Buttons(37).value = tbrPressed
                 End If
             
               ce& = 0 'reset blinker flag
               If GDMDIform.CenterPointTimer.Enabled = True Then
                  ce& = 1 'flag that timer has been shut down during drag
                  GDMDIform.CenterPointTimer.Enabled = False
                  End If
    
               'load previously recorded digitizing results
               ier = ReDrawMap(0)
               If Not InitDigiGraph Then
                  InputDigiLogFile 'load up saved digitizing data for the current map sheet
               Else
                  ier = RedrawDigiLog
                  End If
            
               If DigiRS Then
                  Unload GDRSfrm
                  End If
            
               If numDigiContours > 0 Then 'allow for erasing
                  GDMDIform.Toolbar1.Buttons(40).Enabled = True
                  GDMDIform.mnuEraser.Enabled = True
                  GDMDIform.Toolbar1.Buttons(40).value = tbrUnpressed
                  buttonstate&(40) = 0
                  End If
               
               If numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0 Or numDigiErase > 0 And GDMDIform.mnuDigiSweep.Enabled = False Then
                  GDMDIform.Toolbar1.Buttons(41).Enabled = True
                  GDMDIform.mnuDigiSweep.Enabled = True
                  GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
                  buttonstate&(41) = 0
                  End If
                  
               If numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0 Or numDigiErase > 0 Then 'enable point editing
                  GDMDIform.Toolbar1.Buttons(42).Enabled = True
                  GDMDIform.Toolbar1.Buttons(42).value = tbrUnpressed
                  buttonstate&(42) = 0
                  End If
               
                'this in and out call to tracecontours8 fixes some sort of bug
                Call tracecontours8(GDform1.Picture2, INIT_VALUE, 9)
                
                If Not DigitizePadVis Then
                  GDDigitizerfrm.Visible = True
                  BringWindowToTop (GDDigitizerfrm.hwnd)
                Else
                  BringWindowToTop (GDDigitizerfrm.hwnd)
                  If DigiRightButtonIndex >= 9 Or DigiRightButtonIndex <= 11 Then
                     DigiBackground = &HC0FFFF    'neutral
                     GDDigitizerfrm.lblFunction.Enabled = True
                     GDDigitizerfrm.lblFunction.BackColor = DigiBackground
                     End If
                  End If
                  
               End If
            
          Else
             buttonstate&(47) = 0
             GDMDIform.Toolbar1.Buttons(47).value = tbrUnpressed
             Unload TabConSample_VB_Form
             
             If Not TabletControlVis And Not DigitizeOn Then
                ier = ReDrawMap(0)
                End If
                
             If Not DigitizeOn Then
                buttonstate&(37) = 0
                GDMDIform.Toolbar1.Buttons(37).value = tbrUnpressed
                End If
                
'             If DigitizePadVis Then
'                Unload GDDigitizerfrm
'                End If
                
             End If
             
        End If

End Sub

Public Sub mnuSearchActivated_Click()
   'activate search for digitized points in the drag region
    
    If (TopoMap Or GeoMap) And (numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0) Then
    
          If buttonstate&(15) = 0 Then
             buttonstate&(15) = 1
             GDMDIform.Toolbar1.Buttons(15).value = tbrPressed
             SearchDigi = True
             
             If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(12, 1) 'show the right mode in the digitizer form
             
             'disenable other types of drag window operations
             If DigitizeHardy Then
                 DigitizeHardy = False
                 buttonstate&(43) = 0
                 GDMDIform.Toolbar1.Buttons(43).value = tbrUnpressed
                 
                XminC = 0
                YminC = 0
                XmaxC = 0
                YmaxC = 0
                 
                 End If
              
             If buttonstate&(41) = 1 Then
                buttonstate&(41) = 0
                GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
                DigitizerSweep = False
                End If
            
             If buttonstate&(40) = 1 Then
                buttonstate&(40) = 0
                GDform1.Picture2.MousePointer = vbCrosshair
                DigitizerEraser = False
                End If
            
          Else
             buttonstate&(15) = 0
             GDMDIform.Toolbar1.Buttons(15).value = tbrUnpressed
             SearchDigi = False
             
             If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(12, 0) 'show the right mode in the digitizer form
             
             End If
             
        End If
   
End Sub
Public Sub EditDigitizedPoints()
    If (TopoMap Or GeoMap) And (numDigiPoints > 0 Or numDigiLines > 0 Or numDigiContours > 0) Then
    
        If buttonstate&(42) = 0 Then
           buttonstate&(42) = 1
           GDMDIform.Toolbar1.Buttons(42).value = tbrPressed
           DigiEditPoints = True
           
           If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(15, 1) 'show the right mode in the digitizer form
           
           If ImagePointFile Then
              ier = InitDigiPointsImage
              End If
        
          If buttonstate&(15) = 1 Then
             buttonstate&(15) = 0
             GDMDIform.Toolbar1.Buttons(15).value = tbrUnpressed
             End If
             
         'disenable other types of drag window operations
         If DigitizeHardy Then
             DigitizeHardy = False
             buttonstate&(43) = 0
             GDMDIform.Toolbar1.Buttons(43).value = tbrUnpressed
             
            XminC = 0
            YminC = 0
            XmaxC = 0
            YmaxC = 0
             
             End If
          
         If buttonstate&(41) = 1 Then
            buttonstate&(41) = 0
            GDMDIform.Toolbar1.Buttons(41).value = tbrUnpressed
            DigitizerSweep = False
            End If
        
         If buttonstate&(40) = 1 Then
            buttonstate&(40) = 0
            GDform1.Picture2.MousePointer = vbCrosshair
            DigitizerEraser = False
            End If
            
        If buttonstate&(38) = 1 Then
           GDform1.Picture2.MousePointer = vbCrosshair 'restore crosshair cursor
           buttonstate&(38) = 0
           GDMDIform.Toolbar1.Buttons(38).value = tbrUnpressed
           DigitizeExtendGrid = False
           DigiExtendFirstPoint = False
           End If
            
            
         DigitizePoint = False
         DigitizeLine = False
         DigitizeContour = False
         DigitizeHardy = False
         DigiRS = False
         DigitizerEraser = False
         DigitizerSweep = False
         SearchDigi = False
         
'         'clear canvas and only plot the points
'        If numDigiPoints > 0 Then
'
'           ier = ReDrawMap(0)
'
'           For i = 0 To numDigiPoints - 1
'
'              Xpix = DigiPoints(i).X
'              Ypix = DigiPoints(i).Y
'              Zhgt = DigiPoints(i).Z
'
'             'now draw it
'             GDform1.Picture2.Line (Xpix * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom), Ypix * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom))-(Xpix * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom), Ypix * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom)), PointColor& 'TraceColor
'             GDform1.Picture2.Line (Xpix * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom), Ypix * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom))-(Xpix * DigiZoom.LastZoom + Max(2, 2 * DigiZoom.LastZoom), Ypix * DigiZoom.LastZoom - Max(2, 2 * DigiZoom.LastZoom)), PointColor& 'TraceColor
'
'             'write the elevation value if zoomm >= 1
'             If CInt(DigiZoom.LastZoom) >= 1# Then
'                GDform1.Picture2.CurrentX = Xpix * DigiZoom.LastZoom + Max(4, CInt(DigiZoom.LastZoom))
'                GDform1.Picture2.CurrentY = Ypix * DigiZoom.LastZoom
'                GDform1.Picture2.Fontsize = CInt(8 * DigiZoom.LastZoom)
'                GDform1.Picture2.Font = "Ariel"
'                GDform1.Picture2.ForeColor = PointColor&
'                GDform1.Picture2.Print str$(Zhgt)
'                End If
'
'           Next i
'           End If
         
        Else
           buttonstate&(42) = 0
           GDMDIform.Toolbar1.Buttons(42).value = tbrUnpressed
           DigiEditPoints = False
           
            If XpixLast <> -1 And YpixLast <> -1 Then 'erase the last highlighted mark
                
                gdm = GDform1.Picture2.DrawMode
                gdw = GDform1.Picture2.DrawWidth
                
                GDform1.Picture2.DrawMode = 7
                GDform1.Picture2.DrawWidth = 2
            
               'erase last highlight and return
               GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
               GDform1.Picture2.Line (CLng(XpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)))-(CLng(XpixLast * DigiZoom.LastZoom) + CInt(Max(2, 2 * DigiZoom.LastZoom)), CLng(YpixLast * DigiZoom.LastZoom) - CInt(Max(2, 2 * DigiZoom.LastZoom))), HighLightColor 'TraceColor
               
                'restore drawmode
                GDform1.Picture2.DrawMode = gdm
                GDform1.Picture2.DrawWidth = gdw
               
               XpixLast = -1
               YpixLast = -1
               End If
           
           
           If DigitizePadVis Then Call GDDigitizerfrm.ShowModes(15, 0) 'show the right mode in the digitizer form
           
           'restore all the different sorts of digitized points, lines, and contours
           If DigitizeOn Then
              If Not InitDigiGraph Then
                 InputDigiLogFile 'load up saved digitizing data for the current map sheet
              Else
                 ier = RedrawDigiLog
                 End If
              End If
           
           End If
             
        End If

End Sub
Public Sub mnuGPS_Click()
  'activate or deactivate GPS
  
  If ((Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm") Or _
      (Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat")) Then
  
      If GPSconnected Then
         
            Select Case MsgBox("This action will disconnect any communication with your GPS unit!" _
                               & vbCrLf & vbCrLf & "Proceed with disconnecting GPS communication?", _
                                vbExclamation + vbYesNoCancel + vbDefaultButton3, "GPS disconnect")
               Case vbYes
               
                  GPS_off = True 'GPS off button flag set
               
               Case Else
                  
                  Exit Sub
                  
            End Select
            
         
         GPSconnected = False
         buttonstate&(34) = 0
         GDMDIform.Toolbar1.Buttons(34).value = tbrUnpressed
         Unload GPStest
      
      ElseIf Not GPSconnected Then 'reconnect
      
          Dim DeviceTypeNum As Integer
          
          GPS_off = False
    
          DeviceTypeNum = val(GetSetting(App.Title, "Settings", "GPS_device_name"))
          
          If DeviceTypeNum = 0 And Not GPSSetupVis Then
             DeviceType_Init = True 'flag that using gpssetup to determine the baud rate but not to connect
             GPSsetup.Show 'don't let user move on until he takes care of the baud rate
             GPSsetup.Hide
    '         waitime = Timer
    '         Do Until Timer > waitime + 1
    '            DoEvents
    '         Loop
    '         GPSsetup.cmdScan.Enabled = False
    '         GPSsetup.cboCom.Enabled = False
             
            'ask the user
             Control_Num = 7
             TT1.Style = TTBalloon
             TT1.Icon = TTIconInfo
             TT1.Title = "GPS device"
             TT1.TipText = "Please choose your GPS device"
             TT1.PopupOnDemand = True
             TT1.VisibleTime = 6000                                 'After 6 Seconds tooltip will go away
             TT1.CreateToolTip GPSsetup.frmND100.hwnd
             TT1.Show GPSsetup.frmND100.left / Screen.TwipsPerPixelX + 100, GPSsetup.frmND100.Height / Screen.TwipsPerPixelX - 15 '25 '15  '//In Pixel only
             
             GPSsetup.Show vbModal, Me
      
         Else
            GPS_connect
            End If
            
         End If
         
      End If
     
End Sub

Public Function GPS_connect()
   Dim waitime As Single

   On Error GoTo GPS_connect_Error

   Load GPStest
   
   Exit Function
   
GPS_connect_Error:

End Function
Public Sub GPSInitialization()

   'initializes GPS communication

   On Error GoTo GPSInitialization_Error

   If Not Interpolate_Mode% = 3 And Not WakeUp_Computer Then
   
        Select Case MsgBox("Planning on using a GPS to determine the airplane's position?" _
                    & vbCrLf & vbCrLf & "(If you change your mind, use the GPS button.)", _
                    vbYesNo + vbQuestion, "GPS connection initialization")
                             
            Case vbYes
   
                GPS_timer_trials = 0 'first attempt to connect with the following connection values:
                   
                GPSConnectString0 = GetSetting(App.Title, "Settings", "GPS serial-USB connection string")
                GPSConnectString = GPSConnectString0
                If GPSConnectString = sEmpty Then
                   GPSConnectString = "38400,N,8,1" 'default baud rate, parity, data bit, stop bit
                   End If

                GPS_connect
                
            Case vbNo
                
                Exit Sub
                
         End Select
         
    ElseIf Interpolate_Mode% = 3 Then
    
        If ComPort% = 0 Then
           ComPort% = GetSetting(App.Title, "Settings", "GPS serial-USB COM port")
           End If
         
        GPS_timer_trials = 0 'first attempt to connect with the following connection values:
           
        GPSConnectString0 = GetSetting(App.Title, "Settings", "GPS serial-USB connection string")
        GPSConnectString = GPSConnectString0
        If GPSConnectString = sEmpty Then
           GPSConnectString = "38400,N,8,1" 'default baud rate, parity, data bit, stop bit
           End If
        GPS_connect
        End If
                

   On Error GoTo 0
   Exit Sub

GPSInitialization_Error:

    If Err.Number = 13 Then Resume Next
  
End Sub

Public Sub mnuSearchHeights_Click()
   'activate search for highest elevation in the drag region for ITM or degreees lat/long coordinate systems
    
    If heights And (RSMethod0 Or RSMethod1 Or RSMethod2) And _
       ((Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm") Or _
        (Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat")) Then
    
          If buttonstate&(50) = 0 Then
             buttonstate&(50) = 1
             GDMDIform.Toolbar1.Buttons(50).value = tbrPressed
             HeightSearch = True
             
             If buttonstate&(51) = 1 Then
                buttonstate&(51) = 0
                GDMDIform.Toolbar1.Buttons(51).value = tbrUnpressed
                GenerateContours = False
                GDMDIform.combContour.Visible = False
                
                 'erase any contours
                  ier = ReDrawMap(0)
                
                  'clear old clutter
                    
                  If DigitizeOn Then
                     If Not InitDigiGraph Then
                        InputDigiLogFile 'load up saved digitizing data for the current map sheet
                     Else
                        ier = RedrawDigiLog
                        End If
                     End If
                
                End If
            
          Else
             buttonstate&(50) = 0
             GDMDIform.Toolbar1.Buttons(50).value = tbrUnpressed
             HeightSearch = False
             buttonstate&(50) = 0
             End If
             
        End If
        
End Sub
Public Sub mnuContour_Click()
   'activate search for highest elevation in the drag region for ITM or degreees lat/long coordinate systems
    
    If heights And (RSMethod0 Or RSMethod1 Or RSMethod2) And _
       ((Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm") Or _
        (Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat")) Then
    
          If buttonstate&(51) = 0 Then
             buttonstate&(51) = 1
             GDMDIform.Toolbar1.Buttons(51).value = tbrPressed
             GenerateContours = True
             
             If buttonstate&(50) = 1 Then
                buttonstate&(50) = 0
                GDMDIform.Toolbar1.Buttons(50).value = tbrUnpressed
                HeightSearch = False
                End If
                
            GDMDIform.combContour.Visible = True
            If numContours > 0 Then
               GDMDIform.combContour.Text = str(numContours)
            Else
               GDMDIform.combContour.ListIndex = 6 '(10 meters as default) '2 '(3 meters as default)
               End If

          Else
             buttonstate&(51) = 0
             GDMDIform.Toolbar1.Buttons(51).value = tbrUnpressed
             GenerateContours = False
             buttonstate&(52) = 0 'unpress the profiles button and disenable it
             GDMDIform.Toolbar1.Buttons(52).value = tbrUnpressed
             GDMDIform.Toolbar1.Buttons(52).Enabled = False
             
             GDMDIform.combContour.Visible = False
             
             'erase any contours
              ier = ReDrawMap(0)
            
              'clear old clutter
                
              If DigitizeOn Then
                 If Not InitDigiGraph Then
                    InputDigiLogFile 'load up saved digitizing data for the current map sheet
                 Else
                    ier = RedrawDigiLog
                    End If
                 End If
             
             End If
             
        End If
        

End Sub

Public Sub mnuProfile_Click()
   'horizon profiles
   
   Dim ier%
    
    If buttonstate&(52) = 0 Then
       buttonstate&(52) = 1
       GDMDIform.Toolbar1.Buttons(52).value = tbrPressed
       
       Select Case MsgBox("The horizon will be calculated for: " & "X: " & GDMDIform.Text5 & " Y: " & GDMDIform.Text6 & " Z: " & GDMDIform.Text7 _
                          & vbCrLf & "" _
                          & vbCrLf & "Proceed?" _
                          & vbCrLf & "" _
                          & vbCrLf & "(If this is not correct, answer ""Cancel"", and click on the right place." _
                          & vbCrLf & "Afterwards, push the profile button again.)" _
                          & vbCrLf & "" _
                          , vbOKCancel Or vbInformation Or vbDefaultButton1, "Horizon Profile generation")
       
        Case vbOK
             
           'call dll to calculate the profile.
           'the dll reads the xyz coordinate file and outputs view-angle vs. azimuth
           'which is then displayed as an interactive graph similar to the Map&More program
           'if acceptable, it can be merged into existing horizon profiles
           
           'pick eastern or western horizon
           frmMsgBox.MsgCstm "Pick the horizon:" _
                          & vbCrLf & "", _
                          "Horizon", mbQuestion, 1, False, _
                          "Eastern Horizon", "Western Horizon", "Cancel"
        
            Select Case frmMsgBox.g_lBtnClicked
        
               Case 1
                  'Eastern horizon
                  HorizMode% = 1
                
               Case 2
                  'Western horizon
                  HorizMode% = 2
                
               Case 0, 3
                  HorizMode% = 0
                  GDMDIform.Toolbar1.Buttons(52).value = tbrUnpressed
                  buttonstate&(52) = 0
                
          End Select
          
          If HorizMode% > 0 Then 'call the dll after reading the data from the text file and placing it into a data array that will be passed to the dll
             'execute a very fast binary read to dump all the data simultaneously into the data array (source: http://www.tek-tips.com/faqs.cfm?fid=482)
             'if the coordinates are ITM, then they must be converted into latitude and longitude, this will be left for the dll since it will be faster.
             
             ier% = ShowProfile(HorizMode%)

             End If
       
        Case vbCancel
        
           GDMDIform.Toolbar1.Buttons(52).value = tbrUnpressed
           buttonstate&(52) = 0
       
       End Select
         
    Else
       buttonstate&(52) = 0
       GDMDIform.Toolbar1.Buttons(52).value = tbrUnpressed
       buttonstate&(52) = 0
       End If

End Sub

Public Sub mnuSmooth_Click()

   'smooth selected region using Belgier routine
    If buttonstate&(46) = 0 Then
       buttonstate&(46) = 1
       GDMDIform.Toolbar1.Buttons(46).value = tbrPressed
       Belgier_Smoothing = True
    Else
       buttonstate&(46) = 0
       GDMDIform.Toolbar1.Buttons(46).value = tbrUnpressed
       Belgier_Smoothing = False
       End If
       
End Sub
