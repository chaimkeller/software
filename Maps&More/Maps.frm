VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Maps 
   BackColor       =   &H8000000C&
   Caption         =   "Maps & More"
   ClientHeight    =   12930
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   17295
   Icon            =   "Maps.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   60
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":05DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":0B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":1060
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":11FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":173C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":1C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":21C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":22BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":23B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":24C6
            Key             =   "SnapShot"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":25D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":26EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":27FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":290E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":2A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":2B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":2C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":2D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":2E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":2F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":308C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":319E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":32B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":35CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":38E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":3BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":3F18
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":4232
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":454C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":4866
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":4B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":4E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":51B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":54CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":57E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":5B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":5E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":6136
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":6450
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":676A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":687C
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":7E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":91E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":9502
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":981C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":D83A
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":EBC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":F56E
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":FC48
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":FF62
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":1027C
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":10C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":10F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":11256
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":11570
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":1188A
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":11BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":11EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Maps.frx":121D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   900
      ButtonWidth     =   820
      ButtonHeight    =   741
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   29
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DTMbut"
            Object.ToolTipText     =   "Read  DTM CD-ROM for determining heights"
            ImageIndex      =   50
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "printbut"
            Object.ToolTipText     =   "Print out the current map (w/obstructions if any)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "3Dexplorerbut"
            Object.ToolTipText     =   "run the 3D explorer program"
            ImageIndex      =   59
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "obstructbut"
            Object.ToolTipText     =   "Open prof (obstruction) file"
            ImageIndex      =   47
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Choose magnifcation percentage of maps"
            Style           =   4
            Object.Width           =   800
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "map400but"
            Object.ToolTipText     =   "Display 1/400 scale maps"
            ImageIndex      =   42
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "map50but"
            Object.ToolTipText     =   "Display 1/50 scale maps"
            ImageIndex      =   48
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "worldmap"
            Object.ToolTipText     =   "Display world topo map"
            ImageIndex      =   46
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "placebut"
            Object.ToolTipText     =   "Find stored place"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "gotobut"
            Object.ToolTipText     =   "goto the inputed coordinates"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "leftbut"
            Object.ToolTipText     =   "Move center of map to the left"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "rightbut"
            Object.ToolTipText     =   "Move center of map to the right"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "downbut"
            Object.ToolTipText     =   "Move center of map down"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "upbut"
            Object.ToolTipText     =   "Move center of map up"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "terrabut"
            Object.ToolTipText     =   "Activate the TerraViewer"
            ImageIndex      =   43
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Animatbut"
            Object.ToolTipText     =   "Follow the terraviewer on the 1:50 topo maps"
            ImageIndex      =   44
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Followbut"
            Object.ToolTipText     =   "View map point on the TerraViewer"
            ImageIndex      =   52
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "travelbut"
            Object.ToolTipText     =   "Open travel file"
            ImageIndex      =   51
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "timerbut"
            Object.ToolTipText     =   "change timer interval"
            ImageIndex      =   53
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "pausebut"
            Object.ToolTipText     =   "Pause travel file"
            ImageIndex      =   54
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "stopbut"
            Object.ToolTipText     =   "Stop travel file"
            ImageIndex      =   57
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "showroutebut"
            Object.ToolTipText     =   "Show route on map"
            ImageIndex      =   58
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "sunrisekey"
            Object.ToolTipText     =   "calculate sunrise hroizon"
            ImageIndex      =   37
            Style           =   1
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "sunsetkey"
            Object.ToolTipText     =   "calculate sunset horizon"
            ImageIndex      =   36
            Style           =   1
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tempbut"
            Object.ToolTipText     =   "Average Temperatures"
            ImageIndex      =   60
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1920
         TabIndex        =   19
         Text            =   "100"
         ToolTipText     =   "Map Zoom control"
         Top             =   90
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   12555
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   3263
            MinWidth        =   3263
            Text            =   "For Help, press F1"
            TextSave        =   "For Help, press F1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13229
            MinWidth        =   13229
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "System Resources: OK"
            TextSave        =   "System Resources: OK"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      Height          =   11955
      Left            =   0
      ScaleHeight     =   11895
      ScaleWidth      =   17235
      TabIndex        =   1
      Top             =   510
      Width           =   17295
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   5340
         TabIndex        =   21
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   5160
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer3 
         Interval        =   60000
         Left            =   6000
         Top             =   480
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5160
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin PicClip.PictureClip PictureClip2 
         Left            =   3960
         Top             =   1680
         _ExtentX        =   1296
         _ExtentY        =   1296
         _Version        =   393216
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   11895
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   11920
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
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   7320
            TabIndex        =   9
            Text            =   "0"
            ToolTipText     =   "Goto X coordinate"
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
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   11040
            TabIndex        =   11
            Text            =   "0"
            ToolTipText     =   "Goto elevation"
            Top             =   25
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
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   9240
            TabIndex        =   10
            Text            =   "0"
            ToolTipText     =   "Goto Y coordinate"
            Top             =   25
            Width           =   1240
         End
         Begin VB.TextBox Text4 
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
            ForeColor       =   &H00C00000&
            Height          =   250
            Left            =   5880
            TabIndex        =   8
            Text            =   "0"
            ToolTipText     =   "Depression Angle (degrees)"
            Top             =   25
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text3 
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
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4560
            TabIndex        =   7
            Text            =   "0"
            ToolTipText     =   "Elevation (meters)"
            Top             =   25
            Width           =   735
         End
         Begin VB.TextBox Text2 
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
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   2520
            TabIndex        =   6
            Text            =   "0"
            ToolTipText     =   "Map's Y coordinate"
            Top             =   25
            Width           =   1240
         End
         Begin VB.TextBox Text1 
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
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   600
            TabIndex        =   5
            Text            =   "0"
            ToolTipText     =   "Map's X coordinate"
            Top             =   25
            Width           =   1240
         End
         Begin VB.Label Label7 
            Caption         =   "HGT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   10680
            TabIndex        =   18
            Top             =   60
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "SKYy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8760
            TabIndex        =   17
            Top             =   60
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "SKYx"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6840
            TabIndex        =   16
            Top             =   60
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "DIP"
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
            Height          =   255
            Left            =   5520
            TabIndex        =   15
            Top             =   60
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "HGT(m)"
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
            Height          =   250
            Left            =   3960
            TabIndex        =   14
            Top             =   60
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "ITMy"
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
            Height          =   255
            Left            =   2040
            TabIndex        =   13
            Top             =   60
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "ITMx"
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
            Height          =   135
            Left            =   120
            TabIndex        =   12
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4200
         Top             =   360
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3480
         Top             =   360
      End
      Begin PicClip.PictureClip PictureClip1 
         Index           =   8
         Left            =   960
         Top             =   1200
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin PicClip.PictureClip PictureClip1 
         Index           =   7
         Left            =   960
         Top             =   1800
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin PicClip.PictureClip PictureClip1 
         Index           =   6
         Left            =   1680
         Top             =   1800
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin PicClip.PictureClip PictureClip1 
         Index           =   5
         Left            =   2400
         Top             =   1800
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin PicClip.PictureClip PictureClip1 
         Index           =   4
         Left            =   2400
         Top             =   1200
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin PicClip.PictureClip PictureClip1 
         Index           =   0
         Left            =   1680
         Top             =   1200
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin PicClip.PictureClip PictureClip1 
         Index           =   3
         Left            =   2400
         Top             =   600
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin PicClip.PictureClip PictureClip1 
         Index           =   2
         Left            =   1680
         Top             =   600
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin PicClip.PictureClip PictureClip1 
         Index           =   1
         Left            =   960
         Top             =   600
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         ScaleHeight     =   435
         ScaleWidth      =   675
         TabIndex        =   3
         Top             =   6840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         ScaleHeight     =   435
         ScaleWidth      =   675
         TabIndex        =   2
         Top             =   6840
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu openbatfm 
         Caption         =   "&Open (bat file)"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import DTM segment"
      End
      Begin VB.Menu mnuScanlist 
         Caption         =   "Reload &Scanlist.txt"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu resetoriginfm 
         Caption         =   "&Reset origin"
         Enabled         =   0   'False
      End
      Begin VB.Menu filespace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu fmView 
      Caption         =   "&View"
      Visible         =   0   'False
   End
   Begin VB.Menu fmWindows 
      Caption         =   "&Windows"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu Loadfm 
      Caption         =   "&Load"
      Enabled         =   0   'False
      Begin VB.Menu speedfm 
         Caption         =   "&Speed"
         Begin VB.Menu mihr50fm 
            Caption         =   "&1.    50 mi/hr"
         End
         Begin VB.Menu mihr60fm 
            Caption         =   "&2.    60 mi/hr"
         End
         Begin VB.Menu mihr70fm 
            Caption         =   "&3.    70 mi/hr"
         End
         Begin VB.Menu mihr80fm 
            Caption         =   "&4.    80 mi/hr"
         End
         Begin VB.Menu mihr90fm 
            Caption         =   "&5.    90 mi/hr"
         End
         Begin VB.Menu mihr100fm 
            Caption         =   "&6.  100 mi/hr"
         End
         Begin VB.Menu mihr110fm 
            Caption         =   "&7.  110 mi/hr"
         End
         Begin VB.Menu mihr120fm 
            Caption         =   "&8.  120 mi/hr"
         End
         Begin VB.Menu spacefm 
            Caption         =   "-"
         End
         Begin VB.Menu userspeedfm 
            Caption         =   "&New speed"
         End
         Begin VB.Menu spacebarfm 
            Caption         =   "-"
         End
         Begin VB.Menu speeddefaultfm 
            Caption         =   "&Defined by file"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu routefm1 
         Caption         =   "Load &Route"
         Begin VB.Menu routefm 
            Caption         =   "&Pick the route"
         End
         Begin VB.Menu recoverroutefm 
            Caption         =   "&Recover route"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu optionsfm 
      Caption         =   "&Options"
      Begin VB.Menu topofm 
         Caption         =   "&Topo Maps 50000 Scale"
         Enabled         =   0   'False
         Begin VB.Menu map600fm 
            Caption         =   "&600x600 pixels"
            Checked         =   -1  'True
         End
         Begin VB.Menu map1200fm 
            Caption         =   "&1200x1200 pixels"
         End
      End
      Begin VB.Menu topobar 
         Caption         =   "-"
      End
      Begin VB.Menu DTMlimitsfm 
         Caption         =   "&DTM limits"
      End
      Begin VB.Menu mnuGeoCoordinates 
         Caption         =   "&Geoid for geo. coordinates"
         Begin VB.Menu mnuGeoWithoutCorrect 
            Caption         =   "&Clarke 1880"
         End
         Begin VB.Menu mnuGeoWithCorrect 
            Caption         =   "&WGS84 (GPS)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu geoidbar 
         Caption         =   "-"
      End
      Begin VB.Menu diskDTMfm 
         Caption         =   "&Location of DTMs"
      End
      Begin VB.Menu settingsfm 
         Caption         =   "&Origin Settings"
         Begin VB.Menu originfm 
            Caption         =   "&Set present point as origin"
         End
         Begin VB.Menu resetfm 
            Caption         =   "&Each jump resets origin"
            Checked         =   -1  'True
         End
         Begin VB.Menu dontresetfm 
            Caption         =   "&Don't reset origin on jump"
         End
      End
      Begin VB.Menu dragbar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrag 
         Caption         =   "&Drag options"
         Begin VB.Menu mnuDragDisable 
            Caption         =   "&Disenable Mag. drag"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuMagDragEnable 
            Caption         =   "&Enable Mag. drag"
         End
         Begin VB.Menu spacerExcel 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExcelDrag 
            Caption         =   "Enable E&xportl drag"
         End
         Begin VB.Menu spacerTrig 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTrigDrag 
            Caption         =   "Enable &Trig Pnt drag"
         End
         Begin VB.Menu mnuTrigUndo 
            Caption         =   "&Undo last saved trig change"
         End
         Begin VB.Menu spacerGlitch 
            Caption         =   "-"
         End
         Begin VB.Menu mnuColumnFix 
            Caption         =   "Enable &Column Glitch Fix"
         End
         Begin VB.Menu mnuRowfix 
            Caption         =   "Enable &Row Glitch Fix"
         End
      End
      Begin VB.Menu appendbar 
         Caption         =   "-"
      End
      Begin VB.Menu appendfrm 
         Caption         =   "&Append to Travel file"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu pix600fm 
      Caption         =   "&600pixel"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu pix1200fm 
      Caption         =   "&1200pixel"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu importfm 
      Caption         =   "&Import"
      Enabled         =   0   'False
      Begin VB.Menu importmapfm 
         Caption         =   "&Import Map"
         Enabled         =   0   'False
      End
      Begin VB.Menu importcenterfm 
         Caption         =   "&Center Point"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu searchfm 
      Caption         =   "&Search"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuCrossSection 
      Caption         =   "&CrossSection"
      Enabled         =   0   'False
      Begin VB.Menu mnuFirstPoint 
         Caption         =   "&First Point"
      End
      Begin VB.Menu mnuSecondPoint 
         Caption         =   "&Second Point"
      End
   End
   Begin VB.Menu snapshotfm 
      Caption         =   "S&napshot"
   End
   Begin VB.Menu mnuAirPath 
      Caption         =   "&AirPath"
   End
   Begin VB.Menu mnuGPS_init 
      Caption         =   "&GPS"
      Begin VB.Menu mnuGPSsetup 
         Caption         =   "&Set up GPS baud rate and com port"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuGPS 
         Caption         =   "&Connect to GPS"
      End
      Begin VB.Menu mnuGPS_goto 
         Caption         =   "&Go to GPS coordinates"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu fmHelp 
      Caption         =   "&Help"
      Begin VB.Menu fmMMHelp 
         Caption         =   "&Maps && More &Help"
      End
      Begin VB.Menu fmVersion 
         Caption         =   "&About Maps && More"
      End
   End
End
Attribute VB_Name = "Maps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public resizepic2 As Boolean, resizepic3 As Boolean
Public dragx, dragy

Private Sub appendfrm_Click()
  Maps.StatusBar1.Panels(2) = "Choose a travel file to append to"
  ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
  ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
  CommonDialog1.CancelError = True
  If world = False Then
    CommonDialog1.Filter = "Temporay travel files (*.trf)|*.trf|"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.FileName = terradir$ + "\*.trf"
  Else
    CommonDialog1.Filter = "world travel files (*.wtf)|*.wtf|"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.FileName = "c:\dtm\*.wtf"
   End If
  CommonDialog1.ShowOpen
  ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
  ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
  'read the speed, and place the points into the travel arrays
  'then move the maps to the last point and depress the travel
  'button
  Screen.MousePointer = vbHourglass
  appendfile$ = CommonDialog1.FileName
  openfilnum% = FreeFile
  appendtravel = True
  Open appendfile$ For Input As #openfilnum%
  'record the values in the obs array
  If world = False Then
    For i% = 1 To 9
       Line Input #openfilnum%, doclin$
    Next i%
    speedplac% = InStr(doclin$, "Speed =")
    checkspeed = 0
    If speedplac% <> 0 Then speed = Val(Mid$(doclin$, speedplac% + 7, Len(doclin$)))
    Line Input #openfilnum%, doclin$
    travelnum% = 0
    Do Until EOF(openfilnum%)
       Line Input #openfilnum%, doclin$
       skyxposit% = InStr(1, doclin$, " = ") + 3
       skyyposit% = InStr(skyxposit%, doclin$, " ")
       positend% = InStr(skyyposit% + 1, doclin$, " ")
       T1 = Val(Mid$(doclin$, skyxposit%, skyyposit% - skyxposit%))
       T2 = Val(Mid$(doclin$, skyyposit% + 1, positend% - skyyposit% - 1))
       travelnum% = travelnum% + 1
       ReDim Preserve travel(2, travelnum%)
       travel(1, travelnum%) = T1
       travel(2, travelnum%) = T2
    Loop
    Close #openfilnum%
    'invert last point from SKY to ITM and blit maps
    Mode% = 2 'inverse transform from SKY to ITM
    Call ITMSKY(G11, G22, T1, T2, Mode%)
    kmxc = G11: kmyc = G22
  ElseIf world = True Then
      Line Input #openfilnum%, doclin$
      Input #openfilnum%, travelnum%
      Line Input #openfilnum%, doclin$
      Input #savfilnum%, speed
      speed = speed
      ReDim travel(2, travelnum%)
      For i% = 1 To travelnum%
         Input #openfilnum%, j%
         Input #openfilnum%, travel(1, travelnum%)
         Input #openfilnum%, travel(2, travelnum%)
      Next i%
      Close #openfilnum%
      'use last point as lattest map position, and blit maps
      Maps.Text6.Text = travel(2, travelnum%)
      Maps.Text5.Text = travel(1, travelnum%)
      lon = travel(1, travelnum%)
      lat = travel(2, travelnum%)
      If noheights = False Then
        lg = lon
        lt = lat
        Call worldheights(lg, lt, hgt)
        If hgt = -9999 Then hgt = 0
        Maps.Text3.Text = Str$(hgt)
        hgtworld = hgt
        End If
      cirworld = True
      End If

  showroute = True
  tblbuttons(20) = 1
  travelmode = True
  Toolbar1.Buttons(20).value = tbrPressed
  Call blitpictures
  Screen.MousePointer = vbDefault
End Sub

Private Sub importcenterfm_Click()
  'Maps.StatusBar1.Panels(2) = "Move the cursor to the map's true center and then click."
  'determine new fudx, fudy
  If importcenterfm.Checked = True Then
     importcenterfm.Checked = False
     impcenter = False
  Else
     importcenterfm.Checked = True
     impcenter = True
     End If
End Sub

Private Sub importmapfm_Click()
   'ask user where and which map to import (default imports are contained in f:/eroscities)
   'then blit map--details of the map are stored in its associated .map file
   'see below for details of the .map file
   On Error GoTo errhand
   mydir$ = Dir("c:\eroscities\*.*")
   If mydir$ <> sEmpty Then
      mapdir$ = "c:\eroscities"
   Else
      mapdir$ = "f:\eroscities"
      End If
   ChDir mapdir$
   CommonDialog1.CancelError = True
   CommonDialog1.Filter = "City maps as bitmaps (*.bmp)|*.bmp|City maps as gif files (*.gif)|*.gif|City maps as jpegs (*.jpg)|*.jpg|"
   CommonDialog1.FilterIndex = 2
   CommonDialog1.FileName = mapdir$ + "\*.gif"
   CommonDialog1.ShowOpen
   mapfile$ = CommonDialog1.FileName
   mapinfo$ = Mid$(mapfile$, 1, Len(mapfile$) - 3) + "map"
   ext$ = Mid$(mapfile$, Len(mapfile$) - 2, 3)
   For i% = Len(mapfile$) - 4 To 1 Step -1
      CH$ = Mid$(mapfile$, i%, 1)
      If CH$ = "\" Then
         rootname$ = Mid$(mapinfo$, i% + 1, Len(mapfile$) - 4 - i%)
         Exit For
         End If
   Next i%
   myfile$ = Dir(mapinfo$)
   If myfile$ <> sEmpty Then
      'open it and read map information
      mapinfonum% = FreeFile
      Item% = 0
      Open mapinfo$ For Input As #mapinfonum%
      Do Until EOF(mapinfonum%)
         Line Input #mapinfonum%, doclin$
         If doclin$ = "[Format]" Or doclin$ = "[format]" Then
           Line Input #mapinfonum%, doclin$
           If doclin$ <> ext$ Then
               response = MsgBox("Warning, map format not consistent with format recorded in .map file!", vbExclamation + vbOKCancel, "Maps & More")
               If response = vbCancel Then Exit Sub
               End If
           ext$ = doclin$
           Input #mapinfonum%, xpix, ypix
           blank$ = mapdir$ + "\" + "blank" + LTrim$(RTrim$(Str$(xpix))) + "_" + LTrim$(RTrim$(Str$(ypix))) + "." + ext$
           If Dir(blank$) = sEmpty Then
              response = MsgBox("blank picture file: " & blank$ & " not found", vbCritical + vbOKOnly, "Maps & More")
              Exit Sub
              End If
           Item% = Item% + 1
           If Item% = 3 Then Exit Do
         ElseIf doclin$ = "[capital]" Or doclin$ = "[Capital]" Then
           Input #mapinfonum%, xc, yc
           Item% = Item% + 1
           If Item% = 3 Then Exit Do
         ElseIf doclin$ = "[pixel/km]" Or doclin$ = "[Pixel/km]" Then
           Input #mapinfonum%, pixkm
           Item% = Item% + 1
           If Item% = 3 Then Exit Do
           End If
      Loop
      Close #mapinfonum%
      If Item% < 3 Then
         response = MsgBox("The .map file: " & mapinfo$ & " seems to be missing information.  Check it!", vbCritical + vbOKOnly, "Maps & More")
         Exit Sub
         End If
   Else
      response = MsgBox("No map information file (.map) is associated with this file. " _
                        & "Please create such a file then try again.  The content of this file is: " _
                        & "'[Format]'(c.r.)bmp(c.r.)xpixelsize,ypixelsize(c.r.)(c.r.)'[capital]'(c.r.) " _
                        & "x pixel coord of capital, y pixel coord(c.r.)(c.r.)'[pixel/km]'(c.r.)pixels/km," _
                        & vbCritical + vbOKOnly, "Maps & More")
      'this file has the following format (headers must be present, but in any order):

      '[Format]
      'bmp
      '517,574
      '
      '[capital]
      '269,218
      '
      '[pixel/km]
      '10

      Exit Sub
      End If
   'now look for city name in skyworld.sav and read it's coordinates
   filsav% = FreeFile
   found% = 0
   placnam$ = "start"
   Open drivjk$ + "skyworld.sav" For Input As #filsav%
   Do Until EOF(filsav%)
      oldplacnam$ = placnam$
      Input #filsav%, placnam$, itmx, itmy, itmhgt
      If InStr("abcdefghijklmnopqrstuvwxyz", LCase(Mid$(placnam$, 1, 1))) = 0 Then
         response = MsgBox("Error in skyworld.sav detected after entry: " + oldplacnam$, vbCritical + vbOKOnly, "Maps & More")
         Close #filsav%
         Exit Sub
         End If
      If UCase(Mid$(placnam$, 1, Len(rootname$))) = UCase(rootname$) Then
         l2 = itmx
         l1 = itmy
         Maps.Text6.Text = itmy 'latitude
         Maps.Text5.Text = itmx 'longitude
         Call goto_click
         response = MsgBox("Is this a map for the city: " & placnam$ & "? Check the location on the map.", vbQuestion + vbYesNoCancel, "Maps & More")
         If response = vbCancel Then
            Close #filsav%
            Exit Sub
         ElseIf response = vbYes Then
            found% = 1
            Exit Do
            End If
         End If
   Loop
   Close #filsav%
   If found% = 0 Then
      response = MsgBox("City not found in skyworld.sav. Check the spelling or, if necessary, record it's name and coordinates of the map's reference point in skyworld.sav!", vbCritical, "Maps & More")
      End If

   'now calculate all the information needed to blit the file
   'in place of the world map

   'find extent of file in degrees in the x,y directions
   'first find the radius of the ellipsoid of rotation, Re, at that latitude
   Ra = 6378.136
   Rb = 6356.751
   Re = Sqr(1# / ((Cos(itmy * cd) / Ra) ^ 2 + (Sin(itmy * cd) / Rb) ^ 2))
   deglog = (CDbl(xpix) / CDbl(pixkm)) / (Re * cd * Cos(itmy * cd))
   deglat = (CDbl(ypix) / CDbl(pixkm)) / (Re * cd)
   lon = Maps.Text5.Text
   lat = Maps.Text6.Text
   woxorigin = CDbl(itmx) - (CDbl(xc) / CDbl(xpix)) * deglog
   woyorigin = CDbl(itmy) - (1# - (CDbl(yc) / CDbl(ypix))) * deglat

   mapimport = True
   
   fudx = 0
   fudy = 0
   If pixkm = 4.75 And xpix = 656 And ypix = 554 Then
     'standard www.expedia.com world topo maps
     'so set it's center point
     fudx = 6.38508961403232E-02
     fudy = 0.041708027733435
   ElseIf pixkm = 98.360656 And xpix = 1440 And ypix = 855 Then
     fudx = 0.059
     fudy = -0.0136
     End If
   
   If xpix = 517 And ypix = 574 Then
      'fudx = -0.06
      'fudy = 0.014
      End If
      
   If xpix = 1049 And ypix = 1349 Then 'Lakewood-combined-2.jpg map
      fudx = 0.228
      fudy = -0.289
      End If
      
   If xpix = 10201 And ypix = 5489 Then
      fudx = 0.972605   '-0.0338 '0#  '0.9643837835
      fudy = -0.467 '0.0137 '0#  '0.44781015049
      End If
      
   pixwwi = xpix '+ 10
   pixwhi = ypix '+ 10
   printeroffset = 70 'a printer offset- find the source of it!!!!
   sizewx = Screen.TwipsPerPixelX * pixwwi '# twips in half of picture=8850/2
   sizewy = Screen.TwipsPerPixelY * pixwhi '=8850/2
'   mapPictureform.Width = sizewx + 60 '60 is the size (pixels) of the borders
'   mapPictureform.mapPicture.Width = sizewx
'   mapPictureform.Height = sizewy + 60
'   mapPictureform.mapPicture.Height = sizewy

   If mapwi2 > sizewx + 60 Then
      mapPictureform.Width = sizewx + 60 '60 is the size (pixels) of the borders
      mapPictureform.mapPicture.Width = sizewx
      mapwi = mapPictureform.Width
      mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
      End If
   If maphi2 > sizewy + 60 Then
      mapPictureform.mapPicture.Height = sizewy
      maphi = mapPictureform.Height
      mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
      End If
   If world = True Then
      mapxdif = mapxdif + 35
      mapydif = mapydif + 35
      End If
   'If mapPictureform.Width > Screen.Width Then
   '   mapPictureform.Width = Screen.Width - 60
   '   End If
   'If mapPictureform.Height > Screen.Height Then
   '   mapPictureform.Height = Screen.Height - 1900
   '   End If
   world = True
   kmwx = 2 * deglog / sizewx
   kmwy = deglat / sizewy

   Call loadpictures  'load appropriate map tiles into off-screen buffers
   Call blitpictures   'blit desired portions of the off-screen buffers to the screen
   importcenterfm.Enabled = True
   Exit Sub

errhand:
End Sub

Private Sub Combo1_Change()
   On Error GoTo comerr
   If mapPictureform.Visible = True And Combo1.Text <> sEmpty Then
      magnew = Combo1.Text / 100
      If magnew <> mag And magnew >= 1 Then 'And magnew <= 999 Then
         mag = magnew
         Call blitpictures
         End If
      End If
comerr:
End Sub
Private Sub Combo1_Click()
   If mapPictureform.Visible = True Then
      magnew = Combo1.Text / 100
      If magnew <> mag And magnew >= 1 Then
         mag = magnew
         Call blitpictures
         End If
      End If
End Sub
'Private Sub MDImenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Maps.StatusBar1.Panels(2).Text = X
'End Sub


Private Sub diskDTMfm_Click()
   mapdiskDTMfm.Visible = True
   mapdiskDTMfm.SSTab1.Tab = 0
   ret = SetWindowPos(mapdiskDTMfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
End Sub

Private Sub dontresetfm_Click()
  dojump = False
  dontresetfm.Checked = True
  resetfm.Checked = False
End Sub

Private Sub DTMlimitsfm_Click()
   mapLimitsfm.Visible = True
   ret = SetWindowPos(mapLimitsfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
End Sub

Private Sub fmHelp_Click()
    Maps.StatusBar1.Panels(2) = "Help & Version information"
End Sub

Private Sub fmMMHelp_Click()
    Dim lHelp As Long
    Maps.StatusBar1.Panels(2) = "Maps & More Help files"
    Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
    Maps.CommonDialog2.HelpCommand = cdlHelpContents
    Maps.CommonDialog2.ShowHelp
    waitime = Timer
    Do Until Timer > waitime + 5
       DoEvents
    Loop
    lHelp = FindWindow(vbNullString, "Maps & More Help")
    If lHelp > 0 Then
       ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
       End If
End Sub

Private Sub fmVersion_Click()
   Maps.StatusBar1.Panels(2) = "Version Information"
   mapVersionfm.Visible = True
   dx1 = 0
   dy1 = 20
   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
End Sub


Private Sub Loadfm_Click()
   Maps.StatusBar1.Panels(2) = "Select a speed, or load a travel file"
End Sub

Private Sub map1200fm_Click()
  Call map50butsub
   topotype% = 1
   map600fm.Checked = False
   map1200fm.Checked = True
   pixwi = 1172 'size of Eretz Israel bitmaps in pixels
   pixhi = 1172
   pixwwi = 1182 '603 '599 603 '604
   pixwhi = 1182 '602 '598 602 '604
   printeroffset = 70 'a printer offset- find the source of it!!!!
   sizex = Screen.TwipsPerPixelX * pixwi '# twips in half of picture=8850/2
   sizey = Screen.TwipsPerPixelY * pixhi '=8850/2
   sizewx = Screen.TwipsPerPixelX * pixwwi '# twips in half of picture=8850/2
   sizewy = Screen.TwipsPerPixelY * pixwhi '=8850/2
   'km400x = 40000# / sizex 'm/twips=40000/8850
   'km400y = 40000# / sizey '=40000/8850
   km50x = 10000# / sizex   '=5000/8850
   km50y = 10000# / sizey   '=5000/8850
   kmwx = 360# / sizewx
   kmwy = 180# / sizewy
   mapPictureform.Width = sizex + 60 '60 is the size (pixels) of the borders
   mapPictureform.Height = sizey + 60
   mapPictureform.mapPicture.Width = sizex
   mapPictureform.mapPicture.Height = sizey
   If mapPictureform.Width > Screen.Width Then
      mapPictureform.Width = Screen.Width - 60
      End If
   If mapPictureform.Height > Screen.Height Then
      mapPictureform.Height = Screen.Height - 1900
      End If
   mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
   mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
   If world = True Then
      mapxdif = mapxdif + 35
      mapydif = mapydif + 35
      End If
   mapwi = mapPictureform.Width
   maphi = mapPictureform.Height
   kmxc = kmxc + 5785 * km50x '***********
   kmyc = kmyc - 10615 * km50y
   Call map50butsub


End Sub

Private Sub map600fm_Click()
  Call map50butsub
   If topotype% = 1 Then    '************
      kmxc = Fix(kmxc - 5785 * km50x + 0.5)
      kmyc = Fix(kmyc + 10615 * km50y + 0.5)
      End If
   topotype% = 0
   map1200fm.Checked = False
   map600fm.Checked = True
   pixwi = 594 'size of Eretz Israel bitmaps in pixels
   pixhi = 594
   pixwwi = 604 '603 '599 603 '604
   pixwhi = 604 '602 '598 602 '604
   printeroffset = 70 'a printer offset- find the source of it!!!!
'*********************************************************************
   sizex = Screen.TwipsPerPixelX * pixwi '# twips in half of picture=8850/2
   sizey = Screen.TwipsPerPixelY * pixhi '=8850/2
   sizewx = Screen.TwipsPerPixelX * pixwwi '# twips in half of picture=8850/2
   sizewy = Screen.TwipsPerPixelY * pixwhi '=8850/2
   km400x = 40000# / sizex 'm/twips=40000/8850
   km400y = 40000# / sizey '=40000/8850
   km50x = 5000# / sizex   '=5000/8850
   km50y = 5000# / sizey   '=5000/8850
   kmwx = 360# / sizewx
   kmwy = 180# / sizewy
   mapPictureform.Width = sizex + 60 '60 is the size (pixels) of the borders
   mapPictureform.Height = sizey + 60
   mapPictureform.mapPicture.Width = sizex
   mapPictureform.mapPicture.Height = sizey
   mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
   mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
   If world = True Then
      mapxdif = mapxdif + 35
      mapydif = mapydif + 35
      End If
   mapwi = mapPictureform.Width
   maphi = mapPictureform.Height
  Call map50butsub

End Sub

Private Sub MDIForm_Load()
   Dim lTaskBar As Long
   On Error GoTo errorload
   
   'find Windows version
   SysVersions

   'find location of default drives
   numdriv% = Drive1.ListCount
   
   driveletters$ = "cdefghijklmnop"
   s1% = 0: S2% = 0: s3% = 0: s4% = 0: s5% = 0: s6% = 0
   For i% = 1 To numdriv%
'   For i% = 4 To numdriv%
      drivlet$ = Mid$(driveletters$, i%, 1)
      ChDrive drivlet$
      mypath = drivlet$ + ":\" ' Set the path.
      myname = LCase(Dir(mypath, vbDirectory))   ' Retrieve the first entry.
      Do While myname <> sEmpty   ' Start the loop.
         'Ignore the current directory and the encompassing directory.
         If myname <> "." And myname <> ".." Then
            'Use bitwise comparison to make sure MyName is a directory.
            If (GetAttr(mypath & myname) And vbDirectory) = vbDirectory Then
               myname = LCase(myname)
               If s4% = 0 And myname = "jk_c" Then 'And WinVer <> 5 And WinVer <> 261 And myname = "jk" Then
                  s4% = 1: drivjk_c$ = drivlet$ + ":\jk_c\"
               ElseIf s1% = 0 And myname = "jk" Then 'And (WinVer = 5 Or WinVer = 261) And myname = "jk_c" Then
                  s1% = 1: drivjk$ = drivlet$ + ":\jk\"
               ElseIf S2% = 0 And myname = "fordtm" Then
                  S2% = 1: drivfordtm$ = drivlet$ + ":\fordtm\"
               ElseIf s3% = 0 And myname = "cities" Then
                  s3% = 1: drivcities$ = drivlet$ + ":\cities\"
                  defdriv$ = drivlet$
'               ElseIf s3% = 0 And myname = "cities_d" Then
'                  s3% = 1: drivcities$ = drivlet$ + ":\cities_d\"
'                  defdriv$ = drivlet$
               'ElseIf s4% = 0 And myname = "prom" Then
               '   s4% = 1: drivprom$ = drivlet$ + ":\prom\"
               'ElseIf s5% = 0 And myname = "prof" Then
               '   s5% = 1: drivprof$ = drivlet$ + ":\prof\"
               ElseIf s6% = 0 And myname = "dtm" Then
                  s6% = 1: drivdtm$ = drivlet$ + ":\dtm\"
                  End If
               'If s1% = 1 And S2% = 1 And s3% = 1 And s4% = 1 And s5% = 1 And s6% = 1 Then GoTo cdc1
               If s1% = 1 And S2% = 1 And s3% = 1 And s4% = 1 And s6% = 1 Then GoTo cdc1
               End If 'it represents a directory
            End If
         myname = Dir 'Get next entry
      Loop
   Next i%

cdc1: If s1% = 0 Then
         drivjk$ = InputBox("Can't find the ""jk"" directory, please give the full path name below " + _
                  "(e.g., if jk is a subdirectory of c:\program\random\, then input: ""c:\program\random\"" (ncluding the last backslash)")
         If drivjk$ <> sEmpty Then 'check the directory
            myname = Dir(drivjk$, vbDirectory)
            If myname = sEmpty Then
               response = MsgBox("The directory was not found at at the inputed path.  Do you wan't to try inputing it's path again? (inclue the drive letter, as well as the last backslash, e.g., ""c:\program\random\"")", vbCritical + vbYesNo, "Cal Programs")
               If response = vbYes Then
                  GoTo cdc1
               Else
                  GoTo ce10
                  End If
            Else
               s1% = 1
               End If
         ElseIf drivjk$ = sEmpty Then 'user canceled the operation
            GoTo ce10
            End If
         End If

cdc2: If S2% = 0 Then
         drivfordtm$ = InputBox("Can't find the ""fordtm"" directory, please give the full path name below " + _
                  "(e.g., if fordtm is a subdirectory of c:\program\random\, then input: ""c:\program\random\"" (ncluding the last backslash)")
         If drivfordtm$ <> sEmpty Then 'check the directory
            myname = Dir(drivfordtm$, vbDirectory)
            If myname = sEmpty Then
               response = MsgBox("The directory was not found at at the inputed path.  Do you wan't to try inputing it's path again? (inclue the drive letter, as well as the last backslash, e.g., ""c:\program\random\"")", vbCritical + vbYesNo, "Cal Programs")
               If response = vbYes Then
                  GoTo cdc2
               Else
                  GoTo ce10
                  End If
            Else
               S2% = 1
               End If
         ElseIf drivfordtm$ = sEmpty Then 'user canceled the operation
            GoTo ce10
            End If
         End If

cdc3: If s3% = 0 Then
         drivcities$ = InputBox("Can't find the ""cities"" directory, please give the full path name below " + _
                  "(e.g., if cities is a subdirectory of c:\program\random\, then input: ""c:\program\random\"" (ncluding the last backslash)")
         If drivcities$ <> sEmpty Then 'check the directory
            myname = Dir(drivcities$, vbDirectory)
            If myname = sEmpty Then
               response = MsgBox("The directory was not found at at the inputed path.  Do you wan't to try inputing it's path again? (inclue the drive letter, as well as the last backslash, e.g., ""c:\program\random\"")", vbCritical + vbYesNo, "Cal Programs")
               If response = vbYes Then
                  GoTo cdc3
               Else
                  GoTo ce10
                  End If
            Else
               s3% = 1
               End If
         ElseIf drivcities$ = sEmpty Then 'user canceled the operation
            GoTo ce10
            End If
         End If

cdc4: 'If s4% = 0 Then
      '   drivprom$ = InputBox("Can't find the ""prom"" directory, please give the full path name below " + _
      '            "(e.g., if prom is a subdirectory of c:\program\random\, then input: ""c:\program\random\"" (including the last backslash)")
      '   If drivprom$ <> sEmpty Then 'check the directory
      '      myname = Dir(drivcities$, vbDirectory)
      '      If myname = sEmpty Then
      '         response = MsgBox("The directory was not found at at the inputed path.  Do you wan't to try inputing it's path again? (inclue the drive letter, as well as the last backslash, e.g., ""c:\program\random\"")", vbCritical + vbYesNo, "Cal Programs")
      '         If response = vbYes Then
      '            GoTo cdc4
      '         Else
      '            GoTo ce10
      '            End If
      '         End If
      '   ElseIf drivprom$ = sEmpty Then 'user canceled the operation
      '      GoTo ce10
      '      End If
      '   End If

cdc5: 'If s5% = 0 Then
      '   drivprof$ = InputBox("Can't find the ""prof"" directory, please give the full path name below " + _
      '            "(e.g., if prof is a subdirectory of c:\program\random\, then input: ""c:\program\random\"" (ncluding the last backslash)")
      '   If drivprof$ <> sEmpty Then 'check the directory
      '      myname = Dir(drivprof$, vbDirectory)
      '      If myname = sEmpty Then
      '         response = MsgBox("The directory was not found at at the inputed path.  Do you wan't to try inputing it's path again? (inclue the drive letter, as well as the last backslash, e.g., ""c:\program\random\"")", vbCritical + vbYesNo, "Cal Programs")
      '         If response = vbYes Then
      '            GoTo cdc5
      '         Else
      '            GoTo ce10
      '            End If
      '         End If
      '   ElseIf drivprof$ = sEmpty Then 'user canceled the operation
      '      GoTo ce10
      '      End If
      '   End If

cdc6: If s6% = 0 Then
         drivdtm$ = InputBox("Can't find the ""dtm"" directory, please give the full path name below " + _
                  "(e.g., if dtm is a subdirectory of c:\program\random\, then input: ""c:\program\random\"" (ncluding the last backslash)")
         If drivprof$ <> sEmpty Then 'check the directory
            myname = Dir(drivdtm$, vbDirectory)
            If myname = sEmpty Then
               response = MsgBox("The directory was not found at at the inputed path.  Do you wan't to try inputing it's path again? (inclue the drive letter, as well as the last backslash, e.g., ""c:\program\random\"")", vbCritical + vbYesNo, "Cal Programs")
               If response = vbYes Then
                  GoTo cdc6
               Else
                  GoTo ce10
                  End If
            Else
               s6% = 1
               End If
         ElseIf drivdtm$ = sEmpty Then 'user canceled the operation
            GoTo ce10
            End If
         End If

'      If s1% = 1 And S2% = 1 And s3% = 1 And s4% = 1 And s5% = 1 And s6% = 1 Then GoTo 5
      If s1% = 1 And S2% = 1 And s3% = 1 And s6% = 1 Then GoTo 5
   'if got here, means that couldn't find the cities directory
ce10: MsgBox "Can't locate necessary directories! ABORTING program...Sorry", vbCritical + vbOKOnly, "Cal Programs"
      Call MDIform_queryunload(i%, j%)

   'determine if DTMs are present, and where they are
5  XDIM = 8.33333333333333E-03 'GTOPO30 is default DTM
   YDIM = 8.33333333333333E-03

   myfile = Dir(drivjk$ + "mapcdinfo.sav")
   If myfile = sEmpty Then
      israeldtmcdf = True
      israeldtmcd = israeldtmcdf
      worlddtmcdf = True
      worlddtmcd = worlddtmcdf
      israeldtmf = "j"
      israeldtm = israeldtmf
      worlddtmf = "j"
      worlddtm = worlddtmf
      RdHalYes = False
   Else
      mapinfonum% = FreeFile
      Open drivjk$ + "mapcdinfo.sav" For Input As #mapinfonum%
      Input #mapinfonum%, israeldtmf, israeldtmcdnumf
      israeldtm = israeldtmf
      If israeldtmcdnumf = 0 Then
         israeldtmcdf = False
         israeldtmcd = False
      Else
         israeldtmcdf = True
         israeldtmcd = True
         End If
      Input #mapinfonum%, worlddtmf, worlddtmcdnumf
      worlddtm = worlddtmf
      If worlddtmcdnum = 0 Then
         worlddtmcdf = False
         worlddtmcd = False
      Else
         worlddtmcdf = True
         worlddtmcd = True
         End If
      Input #mapinfonum%, ramdrivef
      ramdrive = ramdrivef
      Input #mapinfonum%, terradirf$
      terradir$ = terradirf$
      Input #mapinfonum%, adx1f, bdy1f
      adx1 = adx1f: bdy1 = bdy1f
      Input #mapinfonum%, RdHalYes
      Close #mapinfonum%
      End If

  If WinVer = 5 Or WinVer = 261 Then 'Windows 2000 or XP
     ramdrive = "e"
     ramdrivef = "e"
     Timer3.Enabled = False 'don't check system resources
     End If
     
  'If WinVer = 261 Then
  '   adx1 = 0.9 * adx1
  '   bdy1 = 0.9 * bdy1
  '   End If
  
  'set Window's taskbar to AutoHide
  'check if this is a reapperance after a reboot
   lnkFile = "c:\windows\startm~1\programs\startup\maps&m~1.lnk"
   myfile = Dir(lnkFile)
   If myfile <> sEmpty Then
      Kill lnkFile
      response = MsgBox("Do you wan't to reenter Maps & More?", vbYesNo + vbQuestion, "Windows")
      If response = vbNo Then End 'Call MDIform_queryunload(i%, j%)
      waitime = Timer
      Do Until Timer > waitime + 3
         DoEvents
      Loop
      reboot = True
      End If

   'place this window above the taskbar
   Me.Left = -60
   Me.Top = -60
   Me.Width = Screen.Width + 120
   Me.Height = Screen.Height + 120
   GoTo map50
   
   'skip this old way of making taskbar disappear
   
'   Screen.MousePointer = vbDefault
'   dx1 = 0
'   dy1 = 1500 '300
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   waitime = Timer  '<<<<<
'   Do Until Timer > waitime + 0.1
'      DoEvents
'   Loop
'   Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'   Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0) 'move mouse to Location item
'   dx1 = 20
'   dy1 = -5
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   waitime = Timer
'   Do Until Timer > waitime + 1 '0.1 '<<<<<<<<<
'      DoEvents
'   Loop
'   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
'   dx1 = 0
'   dy1 = -82
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   'waitime = Timer  '<<<<<
'   'Do Until Timer > waitime + 0.1
'   '   DoEvents
'   'Loop
'   dx1 = -1500
'   dy1 = 0
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   'waitime = Timer  '<<<<<
'   'Do Until Timer > waitime + 3
'   '   DoEvents
'   'Loop
'   dx1 = 30
'   dy1 = 0
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   If reboot = True Then
'      waitime = Timer
'      Do Until Timer > waitime + 1
'         DoEvents
'      Loop
'      reboot = False
'   Else
'      waitime = Timer
'      Do Until Timer > waitime + 1 '0.1 '<<<<<<<<
'         DoEvents
'      Loop
'      End If
'   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
'   waitime = Timer  '<<<<<
'   Do Until Timer > waitime + 0.1
'      DoEvents
'   Loop
'   dx1 = 50
'   dy1 = 60
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   If reboot = True Then
'      waitime = Timer
'      Do Until Timer > waitime + 1
'         DoEvents
'      Loop
'      reboot = False
'      End If
'   'waitime = Timer  '<<<<<
'   'Do Until Timer > waitime + 3
'   '   DoEvents
'   'Loop
'   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
'
'
'
'   'Call keybd_event(VK_RETURN, 0, 0, 0)
'   'Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
'   waitime = Timer
'   Do Until Timer > waitime + 1
'      DoEvents
'   Loop
'
'
'
map50:
   If reboot = True Then
      waitime = Timer
      Do Until Timer > waitime + 1
         DoEvents
      Loop
      reboot = False
      End If
   mapsplash.Visible = True
   waitime = Timer
   Do Until Timer > waitime + 1 'give some time for splashscreen to
                                  'become fully visible for checked for
                                  'the DTM
      DoEvents
   Loop


   'Picture2.MousePointer = vbCrosshair
   'Picture3.MousePointer = vbCrosshair
   'Picture1.MousePointer = 0

'*********************************************************************
'  graphics constants depend on screen size of the 8-bit color bmp's,
'  each having the dimension of = pixwi (pixels) BY pixhi (pixels)
   topotype% = 0
   pixwi = 594 'size of Eretz Israel bitmaps in pixels
   pixhi = 594
   pixwwi = 604 '603 '599 603 '604
   pixwhi = 604 '602 '598 602 '604
   printeroffset = 70 'a printer offset- find the source of it!!!!
'*********************************************************************
   sizex = Screen.TwipsPerPixelX * pixwi '# twips in half of picture=8850/2
   sizey = Screen.TwipsPerPixelY * pixhi '=8850/2
   sizewx = Screen.TwipsPerPixelX * pixwwi '# twips in half of picture=8850/2
   sizewy = Screen.TwipsPerPixelY * pixwhi '=8850/2
   km400x = 40000# / sizex 'm/twips=40000/8850
   km400y = 40000# / sizey '=40000/8850
   km50x = 5000# / sizex   '=5000/8850
   km50y = 5000# / sizey   '=5000/8850
   kmwx = 360# / sizewx
   kmwy = 180# / sizewy
   mapPictureform.Width = sizex + 60 '60 is the size (pixels) of the borders
   mapPictureform.Height = sizey + 60
   mapPictureform.mapPicture.Width = sizex
   mapPictureform.mapPicture.Height = sizey
   mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
   mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
   If world = True Then
      mapxdif = mapxdif + 35
      mapydif = mapydif + 35
      End If
   coordmode% = 1
   speed = 80
   speedmodify = False
   ret = SetWindowPos(mapsplash.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   Maps.MousePointer = 0
   mapwi = mapPictureform.Width
   maphi = mapPictureform.Height
   mag = 1#
   For i% = 0 To 20
      Combo1.AddItem LTrim$(RTrim$(Str(100 + i% * 20)))
   Next i%
   'Maps.Icon = ImageList1.ListImages(30).Picture
   CHMNEO = "XX"
   kmxoo = 0#: kmyoo = 0#
   noheights = False
   hgtTrig = -9999
10 myfile = Dir(israeldtm + ":\dtm\dtm-map.loc")
   If myfile = sEmpty Then
      myfile = Dir(worlddtm + ":\Gt30dem.gif")
      If myfile = sEmpty Then
         noheights = True
         mnuTrigDrag.Enabled = False
         mnuTrigUndo.Enabled = False
         'response = MsgBox("DTM CD not found, please load it into the CD drive. (Pressing the " + _
         '           "CANCEL button will enter SkyLight " + _
         '           "without the option to determine heights.)", vbExclamation + vbOKCancel, "SkyLight")
         'If response = vbOK Then
         '   GoTo 10
         'ElseIf response = vbCancel Then
         '   noheights = True
         '   End If
      Else
         world = True
         noheights = False
         tblbuttons%(1) = 1
         Toolbar1.Buttons(1).value = tbrPressed
         End If
      End If
15 If noheights = False And world = False Then
      Toolbar1.Buttons(1).value = tbrPressed
      tblbuttons%(1) = 1
      filnum% = FreeFile
      Open israeldtm + ":\dtm\dtm-map.loc" For Input As #filnum%
      For i% = 1 To 3
         Line Input #filnum%, doclin$
      Next i%
      N% = 0
      For i% = 4 To 54
         Line Input #filnum%, doclin$
         If i% Mod 2 = 0 Then
            N% = N% + 1
            For j% = 1 To 14
               CHMAP(j%, N%) = Mid$(doclin$, 6 + (j% - 1) * 5, 2)
            Next j%
            End If
      Next i%
      Close #filnum%
      End If


   'load in CD # for USGS EROS DEM (tiles are numbered from
   'left to right, top to bottom - see Cds.gif file)
   worldCD%(1) = 1
   worldCD%(2) = 1
   worldCD%(3) = 1
   worldCD%(4) = 1
   worldCD%(5) = 3
   worldCD%(6) = 3
   worldCD%(7) = 3
   worldCD%(8) = 3
   worldCD%(9) = 3
   worldCD%(10) = 1
   worldCD%(11) = 1
   worldCD%(12) = 1
   worldCD%(13) = 2
   worldCD%(14) = 2
   worldCD%(15) = 2
   worldCD%(16) = 3
   worldCD%(17) = 3
   worldCD%(18) = 4
   worldCD%(19) = 4
   worldCD%(20) = 4
   worldCD%(21) = 2
   worldCD%(22) = 2
   worldCD%(23) = 2
   worldCD%(24) = 2
   worldCD%(25) = 4
   worldCD%(26) = 4
   worldCD%(27) = 4
   worldCD%(28) = 5

   'initial map position
   dojump = True
   myfile = Dir(drivjk$ + "mapposition.sav")
   If myfile = sEmpty Then
      kmxc = 172352 ' 160000 '200000 '172355 '170000 '172355 '170500 '172355
      kmyc = 1131700 '1200000 '1131694 '1130000 '1131694 '1130000 '1131694
      hgtpos = 740.1
      hgt50c = hgtpos
      hgt400c = hgtpos
      kmxsky = kmxc: kmysky = kmyc
      lon = 35.2385
      lat = 31.805042
      hgtworld = 762
      maxangf% = 80
      diflogf% = 4
      diflatf% = 4
      fullrangef% = 0
      viewmode = 3
      modeval = 2.8
      DTMflag = 0
      maxangfs% = 45
      diflogfs% = 2
      diflatfs% = 3
      fullrangefs% = 2
      viewmodes = 3
      modevals = 2.8
      CalculateProfile = 0
      AziStepf% = 10
      rderos2_use = False
      autoazirange% = 1
      TemperatureModel% = 1
   Else
      filnum% = FreeFile
      hgtpos = 0: hgtworld = 0
      Open drivjk$ + "mapposition.sav" For Input As #filnum%
      If kmxc < 70000 Or kmxc > 400000 Or kmyc < 70000 Or kmyc > 1400000 Then
         kmxc = 172352
         kmyc = 1131700
         End If
      Input #filnum%, kmxc, kmyc, hgtpos
      Input #filnum%, lon, lat, hgtworld
      Input #filnum%, maxangf%, diflogf%, diflatf%, fullrangef%, viewmodef%, modevalf
      Input #filnum%, DTMflag
      Input #filnum%, maxangfs%, diflogfs%, diflatfs%, fullrangefs%, viewmodefs%, modevalfs
      Input #filnum%, CalculateProfile
      Input #filnum%, AziStepf%
      Input #filnum%, rderos2_use
      Input #filnum%, IgnoreTiles%
      Input #filnum%, autoazirange%
      Input #filnum%, TemperatureModel%
      Close #filnum%
      
      If Dir(drivjk$ & "mapSRTMinfo.sav") <> sEmpty Then
        mapinfonum% = FreeFile
        Open drivjk$ & "mapSRTMinfo.sav" For Input As #mapinfonum%
        Input #mapinfonum%, srtmdtm, srtmdtmcdnum
        If srtmdtmcdnum = 1 Then
           srtmdtmcd = True
        Else
           srtmdtmcd = False
           End If
        If Dir(srtmdtm & ":\3AS\", vbDirectory) <> sEmpty Or _
           Dir(srtmdtm & ":\USA\", vbDirectory) <> sEmpty Then
           world = True
           noheights = False
           tblbuttons%(1) = 1
           Toolbar1.Buttons(1).value = tbrPressed
           End If
        End If

'      If hgtpos = sEmpty Then hgtpos = 0
'      If hgtworld = sEmpty Then hgtworld = 0
      If hgtpos = 0 And noheights = False And world = False Then
           kmxo = kmxc: kmyo = kmyc
           Call heights(kmxo, kmyo, hgt)
           hgtpos = hgt
           End If
      hgt50c = hgtpos
      hgt400c = hgtpos
      kmxsky = kmxc
      kmysky = kmyc
      End If

  'add maplimit.sav file handler
   maxang% = maxangf%
   diflog% = diflogf%
   diflat% = diflatf%
   viewmode% = viewmodef%
   modeval = modevalf
   fullrange% = fullrangef%
   maxangs% = maxangfs%
   diflogs% = diflogfs%
   diflats% = diflatfs%
   viewmodes% = viewmodefs%
   modevals = modevalfs
   fullranges% = fullrangefs%
   AziStep% = AziStepf%
   noVoidflag = 0 'default is not to smooth out SRTM radar shadows/voids
   
  'Default: subtract 4.7 arc section correction to latitutes and longitudes
   ggpscorrection = True

   waitime = Timer
   If mapsplash.Visible = False Then Exit Sub
   Do Until Timer > waitime + 1
      DoEvents
   Loop
   init = True

   Maps.Visible = True
   'cx = GetSystemMetrics(SM_CXSCREEN)
   'cy = GetSystemMetrics(SM_CYSCREEN)
   'ret = SetWindowPos(Maps.hWnd, HWND_TOP, 0, 0, cx, cy, SWP_SHOWWINDOW)

   waitime = Timer
   Do Until Timer > waitime + 1
      DoEvents
   Loop
   mapsplash.Visible = False
   Unload mapsplash
   mapcapold$ = Maps.Caption
   
   If Dir(drivjk$ & "scanlist.txt") <> sEmpty Then
      mnuScanlist.Enabled = True 'enable opening of scanlist
      End If
      
   Exit Sub

errorload:
   Unload mapsplash
   If filnum% > 0 Then Close #filnum%
   If Err.Number = 71 Then
      response = MsgBox("Drive not ready, try again?", vbCritical + vbOKCancel, "SkyLight")
      If response = vbOK Then
         Resume
      Else
         Call MDIform_queryunload(i%, j%)
         End If
    ElseIf Err.Number = 52 Or Err.Number = 53 Or Err.Number = 75 Or Err.Number = 76 Then
       'response = MsgBox("DTM CD not found, please load it into the CD drive. (Pressing the " + _
       '                  "CANCEL button will enter SkyLight " + _
       '                  "without the option to determine heights.)", vbInformation + vbOKCancel, "SkyLight")
       'If response = vbOK Then
       '   Resume
       'Else
          noheights = True
          myfile = sEmpty
          Resume Next
          'Call form_QueryUnload(i%, j%)
       '   End If
    ElseIf Err.Number = 13 Then
       'bad mapposition.sav file--use defaults
       Resume Next
       End If
       c = Err.Number
End Sub
'Private Sub Form1_MouseUp(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   If Form1.Width <> Form1.Picture1.Width + 150 Then
'      Form1.Picture1.Width = Form1.Width - 150
'      Form1.Picture1.Refresh
'      End If
'   If Form1.Height <> Form1.Picture1.Height + 150 Then
'      Form1.Picture1.Height = Form1.Height - 150
'      Form1.Picture1.Refresh
'      End If
'   End Sub

'   End Sub
'Private Sub MDIForm_DragDrop(Source As Control, _
    X As Single, Y As Single)
'   If Form1.Width <> Form1.Picture1.Width + 150 Then
'      Form1.Picture1.Width = Form1.Width - 150
'      Form1.Picture1.Refresh
'      End If
'   If Form1.Height <> Form1.Picture1.Height + 150 Then
'      Form1.Picture1.Height = Form1.Height - 150
'      Form1.Picture1.Refresh
'      End If
'   End Sub
'_________________________________________________________
'<<>>Private Sub Picture2_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Picture2.Drag 1
'   dragx = X + Picture2.Left + 30
'   dragy = Y + Picture2.Top + 30
'End Sub
'___________________________________________________________
'<<>>Private Sub Picture3_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Picture3.Drag 1
'   dragx = X + Picture3.Left + 30
'   dragy = Y + Picture3.Top + 30
'End Sub
'__________________________________________________________
'<<>>Private Sub Picture1_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Picture3.Drag 1
'   dragx = X
'   dragy = Y
'End Sub
'__________________________________________________________
'<<>>Private Sub Picture1_DragDrop(Source As Control, _
'   X As Single, Y As Single)
'   If TypeOf Source Is PictureBox Then
'      If X - dragx <> 0 Or Y - dragy <> 0 Then
'         If resizepic2 = False And resizepic3 = False Then
'            Source.Move (Source.Left + X - dragx), (Source.Top + Y - dragy)
'         ElseIf resizepic2 = True And resizepic3 = False Then
'            Picture2.Width = Picture2.Width + Abs(X - dragx)
'            Picture2.Height = Picture2.Height + Abs(Y - dragy)
'            If Y - dragy < 0 Then
'               Source.Move (Source.Left + X - dragx), (Source.Top + Y - dragy)
'               End If
'            resizepic2 = False
'         ElseIf resizepic3 = True And resizepic2 = False Then
'            Picture3.Width = Picture3.Width + Abs(X - dragx)
'            Picture3.Height = Picture3.Height + Abs(Y - dragy)
'            If Y - dragy < 0 Then
'               Source.Move (Source.Left + X - dragx), (Source.Top + Y - dragy)
'               End If
'            resizepic3 = False
'            End If
'         End If
'      'Text1.Text = X
'      'Text2.Text = Y
'      'Text3.Text = dragx
'      'Text4.Text = dragy
'      End If
'End Sub
'__________________________________________________
'<<>>Private Sub Picture2_DragDrop(Source As Control, _
'   X As Single, Y As Single)
'   If TypeOf Source Is PictureBox Then
'     If X + Picture2.Left + 30 - dragx <> 0 Or Y + Picture2.Top + 30 - dragy <> 0 Then
'        If resizepic2 = False Then
'           Source.Move (Source.Left + X + Picture2.Left + 30 - dragx), (Source.Top + Y + Picture2.Top + 30 - dragy)
'        ElseIf resizepic2 = True Then
'           Picture2.Width = Picture2.Width - Abs(X + Picture2.Left + 30 - dragx)
'           Picture2.Height = Picture2.Height - Abs(Y + Picture2.Top + 30 - dragy)
'           resizepic2 = False
'           End If
'        End If
'     'Text1.Text = X + Picture2.Left + 30
'     'Text2.Text = Y + Picture2.Top + 30
'     'Text3.Text = dragx
'     'Text4.Text = dragy
'     End If
'End Sub
'_______________________________________________________
'<<>>Private Sub Picture3_DragDrop(Source As Control, _
'   X As Single, Y As Single)
'   If TypeOf Source Is PictureBox Then
'      If X + Picture3.Left + 30 - dragx <> 0 Or Y + Picture3.Top + 30 - dragy <> 0 Then
'         If resizepic3 = False Then
'            Source.Move (Source.Left + X + Picture3.Left + 30 - dragx), (Source.Top + Y + Picture3.Top + 30 - dragy)
'         ElseIf resizepic3 = True Then
'           Picture3.Width = Picture3.Width - Abs(X + Picture3.Left + 30 - dragx)
'           Picture3.Height = Picture3.Height - Abs(Y + Picture3.Top + 30 - dragy)
'           resizepic3 = False
'           End If
'         End If
'      'Text1.Text = X + Picture3.Left + 30
'      'Text2.Text = Y + Picture3.Top + 30
'      'Text3.Text = dragx
'      'Text4.Text = dragy
'      End If
'End Sub
'____________________________________________________
'<<>>Private Sub Picture1_MouseMove(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   'Text1.Text = X
'   'Text2.Text = Y
'   dragx = X
'   dragy = Y
'   'Text3.Text = dragx
'   'Text4.Text = dragy
'End Sub
'_____________________________________________________
'<<>>Private Sub Picture2_MouseMove(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   'Text1.Text = X + Picture2.Left + 30
'   'Text2.Text = Y + Picture2.Top + 30
'   dragx = X + Picture2.Left + 30
'   dragy = Y + Picture2.Top + 30
''   If (X >= 0 And X <= 100) Or _
''      (X >= Picture2.Width - 100 And X <= Picture2.Width + 30) Then
'   If (X >= Picture2.Width - 100 And X <= Picture2.Width + 30) Then
'      Picture2.MousePointer = 9
'      resizepic2 = True
'      Text5.Text = "true"
''   ElseIf (Y >= 0 And Y <= 100) Or _
''      (Y <= Picture2.Height And Y > Picture2.Height - 100) Then
'   ElseIf (Y <= Picture2.Height And Y > Picture2.Height - 100) Then
'      Picture2.MousePointer = 7
'      resizepic2 = True
'      Text5.Text = "true"
'   Else
'      Picture2.MousePointer = vbCrosshair
'      resizepic2 = False
'      Text5.Text = "false"
'      End If
'
'
'   'Text3.Text = dragx
'   'Text4.Text = dragy
'End Sub
'_______________________________________________________
'<<>>Private Sub Picture3_MouseMove(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   'Text1.Text = X + Picture3.Left + 30
'   'Text2.Text = Y + Picture3.Top + 30
'   dragx = X + Picture3.Left + 30
'   dragy = Y + Picture3.Top + 30
''   If (X >= 0 And X <= 100) Or _
''      (X >= Picture3.Width - 100 And X <= Picture3.Width + 30) Then
'   If (X >= Picture3.Width - 100 And X <= Picture3.Width + 30) Then
'      Picture3.MousePointer = 9
'      resizepic3 = True
''   ElseIf (Y >= 0 And Y <= 100) Or _
''      (Y <= Picture3.Height And Y > Picture3.Height - 100) Then
'   ElseIf (Y <= Picture3.Height And Y > Picture3.Height - 100) Then
'      Picture3.MousePointer = 7
'      resizepic3 = True
'   Else
'      Picture3.MousePointer = vbCrosshair
'      resizepic3 = False
'      End If
'   'Text3.Text = dragx
'   'Text4.Text = dragy
'End Sub




Private Sub mihr100fm_Click()
   If mihr100fm.Checked = False Then
      mihr100fm.Checked = True
      speed = 100
      If world = True Then speed = 1000
      mihr50fm.Checked = False
      mihr60fm.Checked = False
      mihr70fm.Checked = False
      mihr80fm.Checked = False
      mihr90fm.Checked = False
      mihr110fm.Checked = False
      mihr120fm.Checked = False
      speeddefaultfm.Checked = False
      userspeedfm.Checked = False
      End If
End Sub

Private Sub mihr110fm_Click()
   If mihr110fm.Checked = False Then
      mihr110fm.Checked = True
      speed = 110
      If world = True Then speed = 1100
      mihr50fm.Checked = False
      mihr60fm.Checked = False
      mihr70fm.Checked = False
      mihr80fm.Checked = False
      mihr90fm.Checked = False
      mihr100fm.Checked = False
      mihr120fm.Checked = False
      speeddefaultfm.Checked = False
      userspeedfm.Checked = False
      End If
End Sub

Private Sub mihr120fm_Click()
   If mihr120fm.Checked = False Then
      mihr120fm.Checked = True
      speed = 120
      If world = True Then speed = 1200
      mihr50fm.Checked = False
      mihr60fm.Checked = False
      mihr70fm.Checked = False
      mihr80fm.Checked = False
      mihr90fm.Checked = False
      mihr100fm.Checked = False
      mihr110fm.Checked = False
      speeddefaultfm.Checked = False
      userspeedfm.Checked = False
      End If
End Sub

Private Sub mihr50fm_Click()
   If mihr50fm.Checked = False Then
      mihr50fm.Checked = True
      speed = 50
      If world = True Then speed = 500
      mihr60fm.Checked = False
      mihr70fm.Checked = False
      mihr80fm.Checked = False
      mihr90fm.Checked = False
      mihr100fm.Checked = False
      mihr110fm.Checked = False
      mihr120fm.Checked = False
      speeddefaultfm.Checked = False
      userspeedfm.Checked = False
      End If
End Sub

Private Sub mihr60fm_Click()
   If mihr60fm.Checked = False Then
      mihr60fm.Checked = True
      speed = 60
      If world = True Then speed = 600
      mihr50fm.Checked = False
      mihr70fm.Checked = False
      mihr80fm.Checked = False
      mihr90fm.Checked = False
      mihr100fm.Checked = False
      mihr110fm.Checked = False
      mihr120fm.Checked = False
      speeddefaultfm.Checked = False
      userspeedfm.Checked = False
      End If
End Sub

Private Sub mihr70fm_Click()
   If mihr70fm.Checked = False Then
      mihr70fm.Checked = True
      speed = 70
      If world = True Then speed = 700
      mihr50fm.Checked = False
      mihr60fm.Checked = False
      mihr80fm.Checked = False
      mihr90fm.Checked = False
      mihr100fm.Checked = False
      mihr110fm.Checked = False
      mihr120fm.Checked = False
      speeddefaultfm.Checked = False
      userspeedfm.Checked = False
      End If
End Sub

Private Sub mihr80fm_Click()
   If mihr80fm.Checked = False Then
      mihr80fm.Checked = True
      speed = 80
      If world = True Then speed = 800
      mihr50fm.Checked = False
      mihr60fm.Checked = False
      mihr70fm.Checked = False
      mihr90fm.Checked = False
      mihr100fm.Checked = False
      mihr110fm.Checked = False
      mihr120fm.Checked = False
      speeddefaultfm.Checked = False
      userspeedfm.Checked = False
      End If
End Sub

Private Sub mihr90fm_Click()
   If mihr90fm.Checked = False Then
      mihr90fm.Checked = True
      speed = 90
      If world = True Then speed = 900
      mihr50fm.Checked = False
      mihr60fm.Checked = False
      mihr70fm.Checked = False
      mihr80fm.Checked = False
      mihr100fm.Checked = False
      mihr110fm.Checked = False
      mihr120fm.Checked = False
      speeddefaultfm.Checked = False
      userspeedfm.Checked = False
      End If
End Sub

Private Sub mnuAirPath_Click()
   If bAirPath Then
      response = MsgBox("Disenable AirPath calculations?.", vbQuestion + vbYesNoCancel, "Maps&More")
      If response = vbYes Then
        bAirPath = False
        If noheights = True Then
           mnuCrossSection.Enabled = False
           mnuFirstPoint.Enabled = False
           mnuSecondPoint.Enabled = False
           End If
        End If
   Else
      bAirPath = True
      mnuCrossSection.Enabled = True
      mnuFirstPoint.Enabled = True
      mnuSecondPoint.Enabled = True
      MsgBox "AirPath calculations enabled." & vbLf & _
      "Use the Cross Section Menu Item to define the" & vbLf & _
      "starting and ending points"
      End If
End Sub

Private Sub mnuColumnFix_Click()
   mnuDragDisable.Checked = False
   mnuMagDragEnable.Checked = False
   mnuExcelDrag.Checked = False
   mnuTrigDrag.Checked = False
   mnuColumnFix.Checked = True
   mnuRowfix.Checked = False
End Sub

Private Sub mnuCrossSection_Click()
   mapCrossSection.Visible = True
End Sub

Private Sub mnuDragDisable_Click()
   'disenable dragging over map to obtain magnification
   'window or to dump 3D data to Excel
   mnuDragDisable.Checked = True
   mnuMagDragEnable.Checked = False
   mnuExcelDrag.Checked = False
   mnuTrigDrag.Checked = False
   mnuColumnFix.Checked = False
   mnuRowfix.Checked = False
End Sub

Private Sub mnuExcelDrag_Click()
   'enable dragging to export 3D data to Excel
   'for plotting in 3D
   mnuDragDisable.Checked = False
   mnuMagDragEnable.Checked = False
   mnuExcelDrag.Checked = True
   mnuTrigDrag.Checked = False
   mnuColumnFix.Checked = False
   mnuRowfix.Checked = False
End Sub

Private Sub mnuExit_Click()
   exit1 = True
   Call MDIform_queryunload(i%, j%)
End Sub

Private Sub mnuFile_Click()
   Maps.StatusBar1.Panels(2) = "Pick one of the files"
End Sub

Private Sub mnuFirstPoint_Click()
   response = MsgBox("If map is positioned at first point (with the correct height), then press ""OK"".  If not, press ""Cancel"" and position map at first point, and then return to this option", vbOKCancel + vbInformation, "Maps & More")
   Select Case response
      Case vbOK
         crosssectionpnt(0, 0) = Val(Maps.Text5.Text)
         crosssectionpnt(0, 1) = Val(Maps.Text6.Text)
         crosssectionhgt(0) = Val(Maps.Text7.Text)
         mapCrossSection.txtlon1.Text = Maps.Text5.Text
         mapCrossSection.txtlat1.Text = Maps.Text6.Text
         mapCrossSection.txthgt1.Text = Maps.Text7.Text
         response = MsgBox("Now position map at second point, and click ""Second Point""", vbOKCancel + vbInformation, "Maps & More")
         mapCrossSection.cmdNext.value = True
      Case vbCancel
         Exit Sub
      Case Else
   End Select
End Sub

Private Sub mnuGeoWithCorrect_Click()
   ggpscorrection = True
   mnuGeoWithoutCorrect.Checked = False
   mnuGeoWithCorrect.Checked = True
End Sub

Private Sub mnuGeoWithoutCorrect_Click()
  ggpscorrection = False
  mnuGeoWithCorrect.Checked = False
  mnuGeoWithoutCorrect.Checked = True
End Sub

Private Sub mnuGPS_Click()

  If mnuGPS.Checked = False Then
  
     'handles GPS enabling and disenabling
      
     GPSInitialization 'initialize communication defaults
      
     Maps.mnuGeoWithoutCorrect.Checked = False
     Maps.mnuGeoWithCorrect.Checked = True
     ggpscorrection = True 'must always use WGS84 geoid with GPS
      
  Else 'disconnect the GPS connect
  
    If GPSconnected Then
           
       Select Case MsgBox("This will disconnect the GPS, proceed?...", _
                           vbYesNoCancel + vbQuestion, "GPS connection")
                           
          Case vbYes
          
             Unload GPStest
'             Maps.GPS_timer.Enabled = False
'             Maps.GPSCom_timer.Enabled = False
             Maps.mnuGPS.Checked = False
             Maps.mnuGPS_goto.Enabled = False
             GPSconnected = False
             
          Case Else
          
             Exit Sub
             
       End Select
        
     ElseIf Not GPSconnected Then  'reconnect
     
        Load GPStest
        
        End If
        
     End If
     
End Sub

Private Sub mnuGPS_goto_Click()
    'goto GPS coordinates
    'if world map, then use the GPS WGS84 coordinates directly to position the map
    'if ITM maps, then first convert from WGS84 lat,lon to ITM
    
    If Not world Then
       Dim N As Long
       Dim E As Long
       Dim lat As Double
       Dim lon As Double
       lat = Val(GPStest.TextLat)
       lon = Val(GPStest.TextLon)
       'convert lat lon to ITM
       Call wgs842ics(lat, lon, N, E)
       kmy = N
       kmx = E
       Maps.Text5.Text = kmx
       Maps.Text6.Text = kmy

    Else
       lato = Val(GPStest.TextLat)
       lono = Val(GPStest.TextLon)
       Maps.Text5.Text = Format(lono, "###0.0#####") '-180# + X * 360# / mappictureform.mappicture.Width
       Maps.Text6.Text = Format(lato, "##0.0#####") '90# - Y * 180# / mappictureform.mappicture.Height
       End If
       
    Screen.MousePointer = vbHourglass
    gotobutton = True
    Call goto_click
    gotobutton = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuGPSsetup_Click()
   GPSsetup.Visible = True
End Sub

Private Sub mnuImport_Click()
  'import DTM segment in xyz format
  'backup DTM, then overwrite
  
  On Error GoTo errhand
  
  'import the file and determine the x,y bounds
   Maps.CommonDialog2.CancelError = True
   Maps.CommonDialog2.Filter = "xyz files (*.xyz)|*.xyz|"
   Maps.CommonDialog2.FilterIndex = 1
   Maps.CommonDialog2.ShowSave

   FileName = Maps.CommonDialog2.FileName
   filnum% = FreeFile
   
   backup% = 0
   response = MsgBox("Backup DTM files before saving changes?" & vbLf & _
                 "(The date will be added as a suffix to the backup tiles)", _
                 vbQuestion + vbYesNoCancel + vbDefaultButton1, "Maps&More")
   If response = vbYes Then
      backup% = 1
      End If
           
    Screen.MousePointer = vbHourglass
    'determine which tile(s) are being used and back them up
    CHFind$ = sEmpty
    Open FileName For Input As #filnum%
    Do Until EOF(filnum%)
        Input #filnum%, kmx, kmy, ztmp
        kmxDTM = kmx * 0.001
        kmyDTM = (kmy - 1000000) * 0.001
        IKMX& = Int((kmxDTM + 20!) * 40!) + 1
        IKMY& = Int((380! - kmyDTM) * 40!) + 1
        NROW% = IKMY&: NCOL% = IKMX&

        'FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE
        Jg% = 1 + Int((NROW% - 2) / 800)
        Ig% = 1 + Int((NCOL% - 2) / 800)
          
        IR% = NROW% - (Jg% - 1) * 800
        IC% = NCOL% - (Ig% - 1) * 800
        IR0% = IR%
        IC0% = IC%
          
        IFN& = (IR% - 1) * 801! + IC%
          
        CHFindTmp$ = CHMAP(Ig%, Jg%)
tp250:  If CHFindTmp$ <> CHFind$ Then
           If filn% > 0 Then Close #filn%
           CHFind$ = CHFindTmp$
             
           If backup% = 1 Then
             'back it up if not already backed up
              FileNew$ = israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
              If Dir(FileNew$) = sEmpty Then
                 FileCopy israeldtm + ":\dtm\" & CHFind$, israeldtm + ":\dtm\" & CHFind$ & "_" & Month(Date) & Day(Date) & Year(Date)
                 End If
           Else
              FileNew$ = israeldtm + ":\dtm\" & CHFind$
              End If
                
           'open the tile for writing
           filn% = FreeFile
           Open FileNew$ For Random As #filn% Len = 2
             
           End If
             

        hgtNew = ztmp
         
        'write the changes to the DTM tile
        'Since roundoff errors in converting from coord to
        'integer indexes, just count columns and rows assuming
        'that the first one has no roundoff error
        IFN& = (IR% - 1) * 801! + IC%
        Put #filn%, IFN&, CInt(hgtNew * 10)
           
    Loop
    Close #filnum%
    Screen.MousePointer = vbDefault
         
    Exit Sub
    
errhand:
   Screen.MousePointer = vbDefault
  
End Sub

Private Sub mnuMagDragEnable_Click()
   'enable dragging to obtain magnification window
   mnuDragDisable.Checked = False
   mnuMagDragEnable.Checked = True
   mnuExcelDrag.Checked = False
   mnuTrigDrag.Checked = False
End Sub

Private Sub mnuRowfix_Click()
   mnuDragDisable.Checked = False
   mnuMagDragEnable.Checked = False
   mnuExcelDrag.Checked = False
   mnuTrigDrag.Checked = False
   mnuColumnFix.Checked = False
   mnuRowfix.Checked = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuScanlist_Click
' DateTime  : 1/27/2004 14:32
' Author    : Chaim Keller
' Purpose   : reloads scanlist.txt for automatic profile analysis
'---------------------------------------------------------------------------------------
'
Private Sub mnuScanlist_Click()

   On Error GoTo mnuScanlist_error
      
   'zero dynamic arrays
   UniqueRoots% = 0
   ReDim FileViewFileName(7, UniqueRoots%)
   ReDim FileViewFileType(UniqueRoots%)
    
   'reload scanlist
   filtmp% = FreeFile
   Open drivjk$ & "scanlist.txt" For Input As #filtmp%
   Do Until EOF(filtmp%)
      
      Line Input #filtmp%, doclin$
      If Trim$(doclin$) = sEmpty Then Exit Do
      If InStr(doclin$, "netz") Then
         nstflg% = 1
         netzskiy$ = "\netz\"
      Else
         nstflg% = 0
         netzskiy$ = "\skiy\"
         End If
      pos% = InStr(doclin$, netzskiy$)
      
      If InStr(doclin$, "eros") <> 0 Then 'world files
         world = True
      Else
         world = False
         End If
      
      lencit% = Len(drivcities$)
      pos1% = InStr(doclin$, drivcities$)
      AbrevDir$ = Mid$(doclin$, pos1% + lencit%, pos% - lencit% - pos1%)
      FileViewDir$ = AbrevDir$
      pos% = InStr(1, doclin$, ".")
      ext$ = Mid$(doclin$, pos%, 4)
      uniqroot$ = Mid$(doclin$, pos% - 8, 8)
      
      doc2$ = drivcities$ & Trim$(AbrevDir$) & netzskiy$
      
      'check if directory really exists
      If Dir(doc2$, vbDirectory) = sEmpty Then
         AbrevDir$ = sEmpty
         response = MsgBox("Directory: " & AbrevDir$ & " doesn't exist!" & vbLf & _
                "Skip this entry?", vbYesNoCancel + vbExclamation, "Maps&More")
         If response = vbYes Then
            GoTo s900
         Else
            Exit Sub
            End If
         End If
       
      proFile$ = doc2$ & uniqroot$ & ext$
      filn$ = uniqroot$ & ext$
      UniqueRoots% = UniqueRoots% + 1
      ReDim Preserve FileViewFileName(7, UniqueRoots%)
      ReDim Preserve FileViewFileType(UniqueRoots%)
      'record file names
      FileViewFileName(0, UniqueRoots% - 1) = filn$
      'record each extension
      FileViewFileName(1, UniqueRoots% - 1) = Mid$(ext$, 2, 3)
      'record each file type (netz or skiy)
      If nstflg% = 1 Then
         FileViewFileType(UniqueRoots% - 1) = 1 'sunrise begins at 1
      Else
        FileViewFileType(UniqueRoots% - 1) = -4 'sunset begins at 4
        End If
s900:
   Loop
   Close #filtmp%

   mapAnalyzefm.Visible = True
   ret = SetWindowPos(mapAnalyzefm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)

Exit Sub

mnuScanlist_error:
    Close
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuScanlist_Click of Form Maps"

End Sub

Private Sub mnuSecondPoint_Click()
   response = MsgBox("If map is positioned at second point (with the correct height), then press ""OK"".  If not, press ""Cancel"" and position map at second point, and then return to this option", vbOKCancel + vbInformation, "Maps & More")
   Select Case response
      Case vbOK
         crosssectionpnt(1, 0) = Val(Maps.Text5.Text)
         crosssectionpnt(1, 1) = Val(Maps.Text6.Text)
         crosssectionhgt(1) = Val(Maps.Text7.Text)
         If crosssectionpnt(1, 0) = crosssectionpnt(0, 0) And _
            crosssectionpnt(1, 1) = crosssectionpnt(0, 1) Then
            response = MsgBox("The second point must be different from the first point! Reposition the map on the second point and then reenter this option", vbOKOnly + vbCritical, "Maps & More")
            mapCrossSection.cmdBack.value = True
            Exit Sub
            End If
         mapCrossSection.txtlon2.Text = Maps.Text5.Text
         mapCrossSection.txtlat2.Text = Maps.Text6.Text
         mapCrossSection.txthgt2.Text = Maps.Text7.Text
         mapCrossSection.cmdNext.value = True
         
         'If world = True Then
         '   response = MsgBox("Do you wan't to follow the nearest path along the earth's surface?", vbYesNoCancel + vbDefaultButton2 + vbQuestion, "Maps & More")
         '   If response = vbYes Then
         '      greatcircle = True
         '   Else
         '      greatcircle = False
         '      End If
         '   End If
         'Call mapCrossSections
      Case vbCancel
         Exit Sub
      Case Else
   End Select
End Sub

Private Sub mnuTrigDrag_Click()
   'enable dragging to add Trig points to DTM
   mnuDragDisable.Checked = False
   mnuMagDragEnable.Checked = False
   mnuExcelDrag.Checked = False
   mnuTrigDrag.Checked = True
   mnuColumnFix.Checked = False
   mnuRowfix.Checked = False
End Sub

Private Sub mnuTrigUndo_Click()
   'undo last trig point fix
   If Dir(israeldtm & ":\dtm\" & CHMNEO & "_" & Month(Date) & Day(Date) & Year(Date)) <> sEmpty Then
      Close
      FileCopy israeldtm & ":\dtm\" & CHMNEO & "_" & Month(Date) & Day(Date) & Year(Date), israeldtm & ":\dtm\" & CHMNEO
      CHMNEO = sEmpty
      End If
  
End Sub

Private Sub openbatfm_Click()
    'pick desired bat file
    resetorigin = False
    On Error GoTo canceldialog
    ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
    ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "bat files (*.bat)|*.bat|"
    CommonDialog1.FilterIndex = 1
    If world = False Then
       CommonDialog1.FileName = drivcities$ + "*.bat"
    Else
       CommonDialog1.FileName = drivcities$ + "eros\*.bat"
       End If
    CommonDialog1.ShowOpen

    FileName = CommonDialog1.FileName
    mapbatlistfm.Text1 = FileName
    mapbatlistfm.Visible = True
    If world = True Then mapbatlistfm.Command2.Enabled = True
    batnum% = FreeFile
    mapbatlistfm.List1.Clear
    Open FileName For Input As #batnum%
    Do Until EOF(batnum%)
       Line Input #batnum%, doclin$
       mapbatlistfm.List1.AddItem doclin$
    Loop
    Close #batnum%
    Exit Sub
canceldialog:
End Sub

Private Sub originfm_Click()
    If world = False Then
       kmxc = kmxoo: kmyc = kmyoo
       kmxsky = kmxc: kmysky = kmyc
       Maps.Text5.Text = kmxc
       Maps.Text6.Text = kmyc
       Maps.Label5.Caption = "ITMx"
       Maps.Label6.Caption = "ITMy"
       coordmode2% = 1
    Else
       If coordmode% = 5 Then
          txt1$ = Maps.Text1.Text
          txt2$ = Maps.Text2.Text
          txt3$ = Maps.Label1.Caption
          txt4$ = Maps.Label2.Caption
          End If
       Maps.Text5.Text = Format(lono, "###0.0#####") '-180# + X * 360# / mappictureform.mappicture.Width
       Maps.Text6.Text = Format(lato, "##0.0#####") '90# - Y * 180# / mappictureform.mappicture.Height
       Maps.Label5.Caption = "long."
       Maps.Label6.Caption = "latit."
       Xworld = X
       Yworld = Y
       cirworld = True
       hgtworld = hgt
       lon = lono
       lat = lato
       Screen.MousePointer = vbHourglass
       Call blitpictures
       Screen.MousePointer = vbDefault
       End If
    If world = True Then
       If coordmode% = 5 Then 'fix some type of timing bug that erases some of the entries
          waitime = Timer
          Do Until Timer > waitime + 0.001
             DoEvents
          Loop
          Maps.Text1.Text = txt1$
          Maps.Text2.Text = txt2$
          Maps.Label1.Caption = txt3$
          Maps.Label2.Caption = txt4$
          End If
       Exit Sub
       End If
    If map400 = True Then
       X400c = X: Y400c = Y
       kmx400c = kmxoo: kmy400c = kmyoo
       kmxc = kmx400c: kmyc = kmy400c
       hgt400c = hgt
       Screen.MousePointer = vbHourglass
       Call blitpictures
       Screen.MousePointer = vbDefault
'       If tblbuttons(19) = 1 Then
'          skyleftjump = True
'          Call skyTERRAgoto
'          End If
    ElseIf map50 = True Then
       'cir50 = True
       X50c = X: Y50c = Y
       kmx50c = kmxoo: kmy50c = kmyoo
       kmxc = kmx50c: kmyc = kmy50c
       hgt50c = hgt
       'now calculate postion on 1:400 map
       Screen.MousePointer = vbHourglass
       Call blitpictures
       Screen.MousePointer = vbDefault
       'cir400 = True
       End If
'       If tblbuttons(19) = 1 Then
'          skyleftjump = True
'          Call skyTERRAgoto
'          End If
End Sub

Private Sub pix1200fm_Click()
  'Call map50butsub
  map1200fm_Click
  'Call map50butsub
End Sub

Private Sub pix600fm_Click()
  'Call map50butsub
  map600fm_Click
  'Call map50butsub
End Sub

Private Sub recoverroutefm_Click()
    Maps.StatusBar1.Panels(2) = "Reload a travel route that crashed during travel"
    If world = False Then Exit Sub

    myfile = Dir(ramdrive + ":\wait.x")
    If myfile <> sEmpty Then
       waitime = Timer
       Do Until Timer > waitime + 0.5
          DoEvents
       Loop
       Exit Sub
       End If

    myfile = Dir(ramdrive + ":\travlog.x")
    If myfile = sEmpty Then Exit Sub
    'recover the travel information
    savfilnum% = FreeFile
    Open ramdrive + ":\travlog.x" For Input As #savfilnum%
    Line Input #savfilnum%, doclin$
    Input #savfilnum%, travelnum%
    Line Input #savfilnum%, doclin$
    Input #savfilnum%, speed
    ReDim travel(2, travelnum%)
    For i% = 1 To travelnum%
      Input #savfilnum%, j%
      Input #savfilnum%, travel(1, j%)
      Input #savfilnum%, travel(2, j%)
    Next i%
    Close #savfilnum%

    'signal Maps & More to consider this as a loaded route
    routeload = True
    showroute = True

    'find last position and go there
    If Dir(ramdrive + ":\lndposit.x") <> sEmpty Then
        numtrys = numtrys + 1
        positfil% = FreeFile
        testfil% = FreeFile
        Open ramdrive + ":\lndposit.x" For Input As #positfil%
        Input #positfil%, skyy, skyx
        Close #positfil%
        lon = skyx: lat = skyy
        'blit to there
        worldmove = True
        Call goto_click
        worldmove = False
        End If

    'reactivate animate timer and play buttons
    Maps.Toolbar1.Buttons(23).Enabled = True
    Maps.Toolbar1.Buttons(24).Enabled = True
    Maps.Toolbar1.Buttons(25).Enabled = True
    tblbuttons(25) = 1
    Maps.Toolbar1.Buttons(25).value = tbrPressed
    Maps.Timer2.Enabled = True

    'now let egg.exe continue moving
    waitime = Timer
    Do Until Timer > waitime + 0.1
       DoEvents
    Loop
    myfile = Dir(ramdrive + ":\mapwait.x")
    If myfile <> sEmpty Then Kill ramdrive + ":\mapwait.x"

End Sub

Private Sub resetfm_Click()
  dojump = True
  dontresetfm.Checked = False
  resetfm.Checked = True
End Sub

Private Sub resetoriginfm_Click()
   mapbatlistfm.Visible = True
   mapbatlistfm.Picture2.Visible = True
   mapbatlistfm.Picture1.Visible = True
   mapbatlistfm.Caption = "Reset the map origin"
   mapbatlistfm.Command2.Enabled = True
   resetorigin = True
   mapbatlistfm.Label4.Caption = "City Name"
   mapbatlistfm.Label2.Caption = "Ref. coord."
   mapbatlistfm.Command3.ToolTipText = "Load Maps & More's center coordinate"
   mapbatlistfm.Text8.Text = sEmpty
   mapbatlistfm.StatusBar1.Panels(1) = "Input coordinates"
End Sub

Private Sub routefm_Click()
  Dim lResult As Long
  On Error GoTo errorroute
  'if obstructions activated-deactivate it
  If tblbuttons(4) = 1 And obstflag = True Then
     tblbuttons(4) = 0
     Maps.Toolbar1.Buttons(4).value = tbrUnpressed
     obstflag = False
     If mapPictureform.Visible = True Then
        Call blitpictures 'erase the obstruction lines
        End If
     Maps.Caption = mapcapold$
     End If
  If world = False Then
          lResult = FindWindow(vbNullString, terranam$)
          If lResult = 0 Or mapPictureform.mapPicture.Visible = False Then
             ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
             response = MsgBox("You must display the maps and activate the terraviewer before loading travel files!", vbOKOnly + vbCritical, "Maps & More")
             ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
             Exit Sub
             End If
          ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          CommonDialog1.CancelError = True
          CommonDialog1.Filter = "Temporay travel files (*.trf)|*.trf|"
          CommonDialog1.FilterIndex = 1
          CommonDialog1.FileName = terradir$ + "\*.trf"
          CommonDialog1.ShowOpen

          'the old way was to:
          'read the file and see if it will fit into one file (up to 200 points)
          'if not, make more files.  The file names are a1.trf,a2.trf,a3.trf,etc.

          dx1 = 0
          dy1 = -100 'move the pointer away from terraviewer window
          Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
          Screen.MousePointer = vbHourglass
          openfile$ = CommonDialog1.FileName
          openfilnum% = FreeFile
          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          myfile = Dir(openfile$)

          Open openfile$ For Input As #openfilnum%
          'record the values in the obs array
          For i% = 1 To 9
             Line Input #openfilnum%, doclin$
          Next i%
          speedplac% = InStr(doclin$, "Speed =")
          checkspeed = 0
          If speedplac% <> 0 Then checkspeed = Val(Mid$(doclin$, speedplac% + 7, Len(doclin$)))
          Line Input #openfilnum%, doclin$
          obsnum% = 0
          Do Until EOF(openfilnum%)
             Line Input #openfilnum%, doclin$
             skyxposit% = InStr(1, doclin$, " = ") + 3
             skyyposit% = InStr(skyxposit%, doclin$, " ")
             positend% = InStr(skyyposit% + 1, doclin$, " ")
             T1 = Val(Mid$(doclin$, skyxposit%, skyyposit% - skyxposit%))
             T2 = Val(Mid$(doclin$, skyyposit% + 1, positend% - skyyposit% - 1))
             Mode% = 2 'inverse transform from SKY to ITM
             Call ITMSKY(G11, G22, T1, T2, Mode%)
             kmxob = G11: kmyob = G22
             obsnum% = obsnum% + 1
             ReDim Preserve obs(2, obsnum%)
             obs(1, obsnum%) = kmxob '* 1000
             obs(2, obsnum%) = kmyob '* 1000 + 1000000
          Loop

         'now write a1.trf files and modify speed if flagged
        '  Seek #openfilnum%, 1
        '  For i% = 1 To 10
        '     Line Input #openfilnum%, doclin$
        '  Next i%
        '  speedposit% = InStr(10, doclin$, "#") - 10
        '  endspeedposit% = InStr(speedposit%, doclin$, " ")
        '  lenspeed% = endspeedposit% - speedposit%
        '  checkspeed = Val(Mid$(doclin$, speedposit%, lenspeed%))
          If checkspeed <> speed And speeddefaultfm.Checked = False Then
             modifyspeed = True
          Else
             modifyspeed = False
             End If
          numPnts% = 1
          'Do Until EOF(openfilnum%)
          '   Line Input #openfilnum%, doclin$
          '   numpnts% = numpnts% + 1
          'Loop
        '  If numpnts% <= travelmax% And modifyspeed = False Then
        '   Close #openfilnum%
          If modifyspeed = False Then
             Close #openfilnum%
             SourceFile = openfile$
             DestinationFile = terradir$ + "\a1.trf"
             FileCopy SourceFile, DestinationFile
             routenum% = 1
             routnum% = 0
             routeX = 0
             routeY = 0
             routeload = True
          'ElseIf numpnts% <= travelmax% And modifyspeed = True Then
          ElseIf modifyspeed = True Then
             Seek #openfilnum%, 1
             savfilnum% = FreeFile
             savfilnam$ = terradir$ + "\a1.trf"
             Open savfilnam$ For Output As #savfilnum%
             For j% = 1 To 8
                Line Input #openfilnum%, doclin$
                Print #savfilnum%, doclin$ 'copy header lines
             Next j%
             Print #savfilnum%, "[POINTS], Speed = " + LTrim$(RTrim$(Str$(speed)))
             Line Input #openfilnum%, doclin$
             Line Input #openfilnum%, doclin$
             Print #savfilnum%, doclin$ 'copy header lines

             Do Until EOF(openfilnum%)
                Line Input #openfilnum%, doclin$
                speedposit% = InStr(10, doclin$, "#") - 11
                speedposit% = InStr(speedposit%, doclin$, " ") + 1
                endspeedposit% = InStr(10, doclin$, "#") - 2
                If checkspeed <> 0 Then
                   newspeed = Val(Mid$(doclin$, speedposit%, endspeedposit - speedposit% + 1)) * (checkspeed / speed)
                Else
                   newspeed = speed
                   End If
                Print #savfilnum%, Mid$(doclin$, 1, speedposit% - 1) + " " + LTrim$(RTrim$(Format(Str$(newspeed), "###0.000"))) + " #NOLABEL#"
             Loop
             Close #openfilnum%
             Close #savfilnum%
             routenum% = 1
             routnum% = 0
             routeX = 0
             routeY = 0
             routeload = True
             End If
        '  Else 'make as many 200 point files as necessary
        '       ' in the new program, this will be never necessary
        '     newpnts% = 0
        '     newfilnum% = 0
        '     routenum% = Int(numpnts% / 200) + 1
        '     'now each latter file includes the last point of the previous file
        '     numpnts% = numpnts% + (routenum% - 1)
        '     routenum% = Int(numpnts% / 200) + 1
        '     routnum% = 0
        'r100:   Seek #openfilnum%, 1 'rewinds the file
        '        newfilnum% = newfilnum% + 1
        '        savfilnum% = FreeFile
        '        savfilnam$ = "e:\terraviewer\a" + LTrim$(RTrim$(Str(newfilnum%))) + ".trf"
        '        Open savfilnam$ For Output As #savfilnum%
        '        Print #savfilnum%, "[HEADER]"
        '        Print #savfilnum%, "ReferHeight = 2"
        '        Print #savfilnum%, "ReferSpeed = 2"
        '        Print #savfilnum%, "HeightAboveGround = 0"
        '        Print #savfilnum%, "ChangePitch = 0"
        '        Print #savfilnum%, "RollWhileTurn = 0"
        '        Print #savfilnum%, "CameraDeltaPitch = 0"
        '        Print #savfilnum%, "[POINTS]"
        '        Print #savfilnum%, "#     X       Z    Height Speed Turn-Accel  Speed-Accel"
        '        If newpnts% = 0 Then
        '           For j% = 1 To 9 + newpnts%
        '              Line Input #openfilnum%, doclin$
        '           Next j%
        '        Else
        '           For j% = 1 To 9 + newpnts% - 1
        '              Line Input #openfilnum%, doclin$
        '           Next j%
        '           End If
        '        maxnum1% = numpnts% - newpnts%
        '        If maxnum1% >= 200 Then
        '           maxnum% = 200
        '        Else
        '           maxnum% = maxnum1%
        '           End If
        '        newerpnts% = 0
        '        For j% = 1 To maxnum%
        '           newpnts% = newpnts% + 1
        '           Line Input #openfilnum%, doclin$
        '           If newpnts% > 200 Then
        '              newerpnts% = newerpnts% + 1
        '              posit% = InStr(1, doclin$, "=")
        '              If modifyspeed = False Then
        '                 Print #savfilnum%, LTrim$(RTrim$(Str(newerpnts% - 1))) + " = " + Mid$(doclin$, posit + 1, Len(doclin$) - posit + 2)
        '              ElseIf modifyspeed = True Then
        '                 newdoc$ = LTrim$(RTrim$(Str(newerpnts% - 1))) + " = " + Mid$(doclin$, posit + 1, Len(doclin$) - posit + 2)
        '                 speedposit% = InStr(10, newdoc$, " 0 ") + 2
        '                 Print #savfilnum%, Mid$(newdoc$, 1, speedposit%) + LTrim$(RTrim$(Str(CInt(speed * 1.6 * (22 / 79))))) + " 50 50"
        '                 End If
        '           Else
        '              If modifyspeed = False Then
        '                 Print #savfilnum%, doclin$
        '              Else
        '                 speedposit% = InStr(10, doclin$, " 0 ") + 2
        '                 Print #savfilnum%, Mid$(doclin$, 1, speedposit%) + LTrim$(RTrim$(Str(CInt(speed * 1.6 * (22 / 79))))) + " 50 50"
        '                 End If
        '              End If
        '        Next j%
        '        Close #savfilnum%
        '        If numpnts% - newpnts% <= 0 Then
        '           Close #openfilnum%
        '           GoTo r200 'finished writing files
        '        Else 'write next trf file
        '           GoTo r100
        '           End If
        '        End If
  ElseIf world = True Then
     lResult = FindWindow(vbNullString, "3D Viewer")
     'If lResult = 0 Then 'activate the 3D Viewer
     '   taskID = Shell("c:\samples\vc98\sdk\graphics\directx\egg\debug\egg.exe", vbNormalFocus)
     '   End If
     ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
     If lResult <> 0 Then ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
     CommonDialog1.CancelError = True
     CommonDialog1.Filter = "world travel files (*.wtf)|*.wtf|"
     CommonDialog1.FilterIndex = 1
     CommonDialog1.FileName = "c:\dtm\*.wtf"
     CommonDialog1.ShowOpen
     Screen.MousePointer = vbHourglass
     openfile$ = CommonDialog1.FileName
     openfilnum% = FreeFile
     ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
     If lResult <> 0 Then ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
     Open openfile$ For Input As #openfilnum%
     'check if speed needs to modified
     Line Input #openfilnum%, doclin$
     Input #openfilnum%, travelnum%
     Line Input #openfilnum%, doclin$
     Input #openfilnum%, checkspeed
     Line Input #openfilnum%, doclin$
     Input #openfilnum%, lon 'read initial travel log coordinates
     Input #openfilnum%, lat
     'now load travel array
     ReDim travel(2, travelnum%)
     travel(1, 1) = lon
     travel(2, 1) = lat
     For i% = 2 To travelnum%
         Line Input #openfilnum%, doclin$
         Input #openfilnum%, travel(1, i%)
         Input #openfilnum%, travel(2, i%)
     Next i%
     Close #openfilnum%
     If checkspeed <> speed And speeddefaultfm.Checked = False Then
        modifyspeed = True
     Else
        modifyspeed = False
        End If

     If modifyspeed = False Then
        FileCopy openfile$, ramdrive + ":\travlog.x"
     Else 'write travelog file
        savfilnum% = FreeFile
        Open ramdrive + ":\travlog.x" For Output As #savfilnum%
        Print #savfilnum%, "Number of route points"
        Print #savfilnum%, travelnum%
        Print #savfilnum%, "Speed"
        Print #savfilnum%, speed
        ReDim travel(2, travelnum%)
        For i% = 1 To travelnum%
           Print #savfilnum%, i%
           Print #savfilnum%, travel(1, i%)
           Print #savfilnum%, travel(2, i%)
        Next i%
        Close #savfilnum%
        End If
      waitime = Timer
      Do Until Timer > waitime + 1
         DoEvents
      Loop
      routenum% = 1
      routnum% = 0
      routeX = 0
      routeY = 0
      Screen.MousePointer = vbDefault

     'now goto first point of travelog on maps and check if
     '3D Viewer is activated
      'cirworld = True
      'hgtworld = hgt
      'lono = lon
      'lato = lat
      routeload = True
      showroute = True
      worldmove = True
      Call goto_click
      worldmove = False
      waitime = Timer
      Do Until Timer > waitime + 0.5
         DoEvents
      Loop
      'Screen.MousePointer = vbHourglass
      'Call blitpictures
      'Screen.MousePointer = vbDefault
      'go there on 3D Viewer, if activated
      If lResult <> 0 Then
        'send message to 3D Viewer to begin playing
         nmsg = SendMessage(lResult, WM_COMMAND, 1002, 0)
         GoTo r200
      Else
        'check if there is a USGUS EROS CD in the CD-drive
        On Error GoTo rsunrerr
        myfile = Dir(worlddtm + ":\Gt30dem.gif")
        If myfile = sEmpty Then
           'check if there are stored DTM files in c:\dtm
            doclin$ = Dir("c:\dtm\*.BIN")
            myfile = Dir("c:\dtm\eros.tm3")
            If doclin$ <> sEmpty And myfile <> sEmpty And Dir("c:\dtm\*.BI1") <> sEmpty Then
              'leave rest of checking for sunrisesunset routine
               checkdtm = True
               Call sunrisesunset(1)
            ElseIf Not NoCDWarning Then
               Maps.Toolbar1.Buttons(26).value = tbrUnpressed
               ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
               response = MsgBox("USGS EROS CD not found!  Please enter the appropriate CD, and then press the DTM button!", vbCritical + vbOKOnly, "Maps & More")
               ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
               NoCDWarning = True
               Exit Sub
               End If
        Else
            Call sunrisesunset(1)
            End If
        lResult = FindWindow(vbNullString, "3D Viewer")
        Do Until lResult > 0
           DoEvents
           lResult = FindWindow(vbNullString, "3D Viewer")
        Loop
        End If
     End If

r200:   Screen.MousePointer = vbDefault
        Maps.Toolbar1.Buttons(23).Enabled = True
        Maps.Toolbar1.Buttons(24).Enabled = True
        Maps.Toolbar1.Buttons(25).Enabled = True
'        If world = True Then 'follow the 3D Viewer
'            If tblbuttons(18) = 0 Then
'               tblbuttons(18) = 1
'               Toolbar1.Buttons(18).Value = tbrPressed
'               Maps.Timer2.Enabled = True
'               End If
'           Exit Sub
'           End If

        If world = False Then
           Call routeform
           Do Until insiderouteform = False
              DoEvents
           Loop
           End If
        If routeload = False Then Exit Sub
        If world = True Then routeload = False
        tblbuttons(25) = 1
        Maps.Toolbar1.Buttons(25).value = tbrPressed
        Maps.Timer2.Enabled = True
        showroute = True
        If world = False Then
           Screen.MousePointer = vbHourglass
           Call blitpictures
           Screen.MousePointer = vbDefault
        'ElseIf world = True Then
          'worldmove = True
          'Call goto_click
          'worldmove = False
          'routeload = True '<<<---changed
          End If
        Exit Sub
errorroute:
'  cancel button in Common Dialog Box was pushed
   If world = False Then
        dx1 = 0
        dy1 = -120 'move the pointer away from terraviewer window
        Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
        ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
        ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
   Else
        ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
        ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
        End If
   Exit Sub

rsunrerr:
   myfile = sEmpty
   Resume Next

End Sub

Private Sub routefm1_Click()
  Maps.StatusBar1.Panels(2) = "Pick and load a stored route"
End Sub

Private Sub searchfm_Click()
   If mapsearchfm.Visible = True Then
      ret = SetWindowPos(mapsearchfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
   Else
      mapsearchfm.Visible = True
      OverhWnd = FindWindow(vbNullString, "Overview")
      ret = SetWindowPos(mapsearchfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      'Call BringWindowToTop(OverhWnd)
      ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
      End If
End Sub

Private Sub snapshotfm_Click()
    If world = True Then
       Screen.MousePointer = vbHourglass
       Call keybd_event(VK_SNAPSHOT, 0, 0, 0)
       Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0)
       timwait = Timer
       Do Until Timer > timwait + 0.5
          DoEvents
       Loop
       Screen.MousePointer = vbDefault
       If world = True And mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
       response = MsgBox("The BitMap of the world map has been saved to the Clipboard. " + _
                       "Use MSPaint or an equivalent program to edit/print it.", vbInformation + vbOKOnly, "Maps & More")
       If world = True And mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
       Exit Sub
       End If
    lResult = FindWindow(vbNullString, terranam$)
    If lResult > 0 Then
       Screen.MousePointer = vbHourglass
       ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
       timwait = Timer
       Do Until Timer > timwait + 0.5
         DoEvents
       Loop
       'send a bitmap image of the current window to the CLIPBOARD
       Call keybd_event(VK_SNAPSHOT, 0, 0, 0)
       Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0)
       waittime = Timer + 1
       Do Until Timer > waittime
          DoEvents
       Loop
       Screen.MousePointer = vbDefault
       lResult = FindWindow(vbNullString, terranam$)
       If lResult > 0 Then 'remove topmost status from terraviewer in order to display message
         ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         End If
       If mapPictureform.Visible = True Then 'remove topmost status from map
         ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         End If

       response = MsgBox("The BitMap of the TerraViewere Picture has been saved to the Clipboard. " + _
                  "Use MSPaint or an equivalent program to edit/print it.", vbInformation + vbOKOnly, "Maps & More")
       If lResult > 0 Then 'restore topmost status
          ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
       If mapPictureform.Visible = True Then 'restore topmost status to map
          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
      'ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
      End If

End Sub

Private Sub speeddefaultfm_Click()
      speeddefaultfm.Checked = True
      mihr50fm.Checked = False
      mihr60fm.Checked = False
      mihr70fm.Checked = False
      mihr80fm.Checked = False
      mihr90fm.Checked = False
      mihr100fm.Checked = False
      mihr110fm.Checked = False
      mihr120fm.Checked = False
      userspeedfm.Checked = False
End Sub



Private Sub speedfm_Click()
   Maps.StatusBar1.Panels(2) = "Pick the desired speed, or input desired speed"
End Sub

Private Sub Timer1_Timer()
   If mapimport = False And world = True Then
      deglog = 180
      deglat = 180
      End If
   If mapPictureform.Visible = False Then
      Timer1.Enabled = False
      For i% = 12 To 15
         tblbuttons(i%) = 0
         Toolbar1.Buttons(i%).value = tbrUnpressed
      Next i%
      Exit Sub
      End If
   If map50 = True Then
      Step = 50 / mag
   ElseIf map400 = True Then
      Step = 400 / mag
   ElseIf world = True Then
      'Step = 3.6 / mag
      Step = deglog / (mag * 50)
      End If
   If tblbuttons%(12) = 1 Then
      If world = False Then
         kmxc = kmxc - Step
      Else
         lon = lon - Step
         End If
      Call blitpictures   'blit desired portions of the off-screen buffers to the screen
      GoTo t110
      End If
   If tblbuttons%(13) = 1 Then
      If world = False Then
         kmxc = kmxc + Step
      Else
         lon = lon + Step
         End If
      Call blitpictures   'blit desired portions of the off-screen buffers to the screen
      GoTo t110
      End If
t110:
   If tblbuttons%(14) = 1 Then
      If world = False Then
         kmyc = kmyc - Step
      Else
         lat = lat - Step
         End If
      Call blitpictures   'blit desired portions of the off-screen buffers to the screen
      GoTo t200
      End If
   If tblbuttons%(15) = 1 Then
      If world = False Then
         kmyc = kmyc + Step
      Else
         lat = lat + Step
         End If
      Call blitpictures   'blit desired portions of the off-screen buffers to the screen
      End If
t200:
     Call showcoord
End Sub

Private Sub Timer2_Timer()
  Dim bRtn As Boolean, lResult As Long, C1 As String, C2 As String
  Dim lwin As Long
  If world = False Then
      lResult = FindWindow(vbNullString, terranam$)
      'lwin = FindWindow(vbNullString, "Jump to Location")
      'If lwin <> 0 Then
      '      Exit Sub
      '      End If
      If lResult > 0 Then
        'check if window already activated, if so deactivate it
         lwin = FindWindow(vbNullString, "Jump to Location")
         If lwin <> 0 Then
            ret = SetWindowPos(lwin, HWND_TOPMOST, 0, 0, 0, 0, SWP_HIDEWINDOW)
            Call keybd_event(VK_ESCAPE, 0, 0, 0)
            Call keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)
            End If

            'Maps.StatusBar1.Panels(2) = "In order to follow the TerraExplorer, Place the cursor on it's X: or Y: edit box"
    '       Screen.MouseIcon = LoadPicture("c:/progra~1/devstu~1/Vb/Graphics/Icons/Misc/Timer01.ico")
    '        Screen.MouseIcon = LoadPicture("c:/progra~1/devstu~1/Vb/Graphics/Cursors/C_wai03.cur")
    '        Screen.MousePointer = 99
         'this how it used to done, SOB!!!
         'Skycoord% = 1
         'bRtn = EnumChildWindows(lResult, AddressOf EnumFunc, 0) 'read captions

         'bedieved have to do it differently
         'signal the user to keep pointer over X:, or Y: coordinate
         'box of TerraExplorer program

         '!!!!!found a way to restore some of what used to be done
         Skynum% = 0
         bRtn = EnumChildWindows(lResult, AddressOf EnumFunc, 0) 'read captions
         timerwait = Timer + 0.01
         Do Until Timer > timerwait
           DoEvents
         Loop
         Call PostMessage(hChild, WM_SETFOCUS, 0, 0)
    '*********************************************************************
       'depress left mouse button, then read inputs (but hide "Jump to
       'Location window as much as possible)

       'Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'drop-down Location menu
       'Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
       'timerwait = Timer + 0.0001
       timerwait = Timer + 0.01
       Do Until Timer > timerwait
         DoEvents
       Loop
       lwin = FindWindow(vbNullString, "Jump to Location")
       If lwin = 0 Then Exit Sub
       ret = SetWindowPos(lwin, HWND_TOPMOST, 0, 0, 0, 0, SWP_HIDEWINDOW)

       Call keybd_event(VK_CONTROL, 0, 0, 0) 'enter SKYx
       Call keybd_event(VK_INSERT, 0, 0, 0)
       Call keybd_event(VK_INSERT, 0, KEYEVENTF_KEYUP, 0)
       Call keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
       timerwait = Timer + 0.000001
       Do Until Timer > timerwait
         DoEvents
       Loop
       C1 = Clipboard.GetText(vbCFText)
       Call keybd_event(VK_TAB, 0, 0, 0)
       Call keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0)

       Call keybd_event(VK_CONTROL, 0, 0, 0) 'enters SKYy
       Call keybd_event(VK_INSERT, 0, 0, 0)
       Call keybd_event(VK_INSERT, 0, KEYEVENTF_KEYUP, 0)
       Call keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0) 'enters return and goes to desired position
       timerwait = Timer + 0.000001
       Do Until Timer > timerwait
         DoEvents
       Loop
       C2 = Clipboard.GetText(vbCFText)

       'waitime = Timer
       'Do Until Timer > waitime + 0.01
       '   DoEvents
       'Loop
       ret = SetWindowPos(lwin, HWND_TOPMOST, 0, 0, 0, 0, SWP_HIDEWINDOW)
       Call keybd_event(VK_ESCAPE, 0, 0, 0)
       Call keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)
150:   lwin = FindWindow(vbNullString, "Jump to Location")
       If lwin <> 0 Then
          ret = SetWindowPos(lwin, HWND_TOPMOST, 0, 0, 0, 0, SWP_HIDEWINDOW)
          Call keybd_event(VK_ESCAPE, 0, 0, 0)
          Call keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)
          GoTo 150
          End If

    '*********************************************************************


         skyx = Val(C1)
         skyy = Val(C2)
         If skyx = skyy And C1 = C20 Then Exit Sub
         C10 = C1
         C20 = C2

        'check if coordinates are witihin bounds
        If skyx < 50000 Or skyx > 400000 Then Exit Sub
        If skyy < -200000 Or skyy > 400000 Then Exit Sub


         Maps.Label5.Caption = "SKYx"
         Maps.Label6.Caption = "SKYy"
         coordmode2% = 4

         skymove = True
         Call goto_click
         skymove = False
         If routeload = True Then
            If skyx = 0 And skyy = 0 Then Exit Sub
           'just means that haven't yet put cursor on X:, Y: TerraExplorer boxes

            If skyx <> routeX Or skyy <> routeY Then
               routeX = skyx: routeY = skyy
            Else
               'finished current travel file, load next file
               If routnum% < routenum% Then
                  Call routeform
                  routeX = 0: routeY = 0
                  Maps.Toolbar1.Buttons(23).Enabled = True
                  Maps.Toolbar1.Buttons(24).Enabled = True
                  Maps.Toolbar1.Buttons(25).Enabled = True
                  showroute = False
               Else
                  routeload = False
                  travelmode = False
                  routeX = 0: routeY = 0
                  routenum% = 0
                  routnum% = 0
                  Maps.Timer2.Interval = 500
                  Maps.Toolbar1.Buttons(18).value = tbrUnpressed
                  Maps.Toolbar1.Buttons(23).Enabled = False
                  Maps.Toolbar1.Buttons(24).Enabled = False
                  Maps.Toolbar1.Buttons(17).Enabled = True
                  tblbuttons(18) = 0
                  tblbuttons(25) = 0
                  Maps.Timer2.Enabled = False
                  If showroute = True Then
                     showroute = False
                     Maps.Toolbar1.Buttons(25).value = tbrUnpressed
                     Maps.Toolbar1.Buttons(25).Enabled = False
                     Call blitpictures
                     End If
                  End If
               End If
            End If
    '     reset after animation is over
         'ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
         'ret = BringWindowToTop(lResult) 'bring TerraViewer back to top of Z order
        Maps.StatusBar1.Panels(2) = ""
        Screen.MouseIcon = LoadPicture("")
        Screen.MousePointer = vbDefault
        End If
    ElseIf world = True Then
        'first check to see that the scene has finished rendering
        myfile = Dir(ramdrive + ":\wait.x")
        If myfile <> sEmpty Then
           waitime = Timer
           Do Until Timer > waitime + 0.5
              DoEvents
           Loop
           Exit Sub
           End If

        lResult = FindWindow(vbNullString, "3D Viewer")
        If lResult > 0 Then

          'send message to egg not to rerender or relocate until this operation finishes
          messagfil% = FreeFile
          Open ramdrive + ":\mapwait.x" For Output As #messagfil%
          Print #messagfil%, "Hey, wait for me!"
          Close #messagfil%

         numtrys = 0
         On Error GoTo positerror
         If Dir(ramdrive + ":\lndposit.x") <> sEmpty Then
            'DoEvents
            numtrys = numtrys + 1
            positfil% = FreeFile
            testfil% = FreeFile
            Open ramdrive + ":\lndposit.x" For Input As #positfil%
            Input #positfil%, skyy, skyx
            Close #positfil%
            lon = skyx: lat = skyy
            End If
        Else
          If routeload = True Or showroute = True Then
             routeload = False
             showroute = False
             routeX = 0: routeY = 0
             routenum% = 0
             routnum% = 0
             Maps.Timer2.Enabled = False
             Maps.Timer2.Interval = 1500
             tblbuttons(18) = 0
             Maps.Toolbar1.Buttons(18).value = tbrUnpressed
             For i% = 23 To 27
                Maps.Toolbar1.Buttons(i%).Enabled = False
                Maps.Toolbar1.Buttons(i%).value = tbrUnpressed
                tblbuttons(i%) = 0
             Next i%
             If showroute = True Then
                showroute = False
                Call blitpictures
                End If
             Exit Sub
             End If
           End If

        If routeload = True And lResult > 0 Then
            If skyx <> routeX Or skyy <> routeY Then
               routeX = skyx: routeY = skyy
              'check that readDTM is not currently operating
              '(if blit now, can cause this program to crash)
              'lResult = FindWindow(vbNullString, "Extracting relevant portion of the DTM")
              'If lResult > 0 Then Exit Sub
              'otherwise attempt to read the 3D Viewer position file
              'DoEvents
              worldmove = True
              Call goto_click
              worldmove = False
            Else
               'finished current travel file or pausing
               'look for g:\routefin.x or g:\routepau.x files
               If Dir(ramdrive + ":\routefin.x") <> 0 Then
                  routeload = False
                  routeX = 0: routeY = 0
                  routenum% = 0
                  routnum% = 0
                  Maps.Timer2.Interval = 1500
                  Maps.Toolbar1.Buttons(18).value = tbrUnpressed
                  Maps.Toolbar1.Buttons(23).Enabled = False
                  Maps.Toolbar1.Buttons(24).Enabled = False
                  tblbuttons(18) = 0
                  tblbuttons(25) = 0
                  Maps.Timer2.Enabled = False
                  If showroute = True Then
                     showroute = False
                     Maps.Toolbar1.Buttons(25).value = tbrUnpressed
                     Maps.Toolbar1.Buttons(25).Enabled = False
                     worldmove = True
                     Call goto_click
                     worldmove = False
                     End If
                ElseIf Dir(ramdrive + ":\routepau.x") <> 0 Then
                   tblbuttons(23) = 1
                   Toolbar1.Buttons(23).value = tbrPressed
                Else
                    tblbuttons(23) = 0
                    Toolbar1.Buttons(23).value = tbrUnpressed
                    tblbuttons(24) = 0
                    Toolbar1.Buttons(24).value = tbrUnpressed
                    worldmove = True
                    Call goto_click
                    worldmove = False
                   End If
                End If

             'Maps.Text5.Text = Format(lon, "###0.0#####") '-180# + X * 360# / mappictureform.mappicture.Width
             'Maps.Text6.Text = Format(lat, "##0.0#####") '90# - Y * 180# / mappictureform.mappicture.Height
             'If coordmode% = 5 Then
             '   txt1$ = Maps.Text1.Text
             '   txt2$ = Maps.Text2.Text
             '   txt3$ = Maps.Label1.Caption
             '   txt4$ = Maps.Label2.Caption
             '   End If
             ''Maps.Text5.Text = Format(lon, "###0.0#####") '-180# + X * 360# / mappictureform.mappicture.Width
             ''Maps.Text6.Text = Format(lat, "##0.0#####") '90# - Y * 180# / mappictureform.mappicture.Height
             ''Maps.Label5.Caption = "long."
             ''Maps.Label6.Caption = "latit."
             'Xworld = X
             'Yworld = Y
             'cirworld = True
             'hgtworld = hgt
             ''lon = lono
             ''lat = lato
             'Screen.MousePointer = vbHourglass
             'Call blitpictures
             'Screen.MousePointer = vbDefault
             'If coordmode% = 5 Then 'fix some type of timing bug that erases some of the entries
             '   waitime = Timer
             '   Do Until Timer > waitime + 0.001
             '       DoEvents
             '   Loop
             '   Maps.Text1.Text = txt1$
             '   Maps.Text2.Text = txt2$
             '   Maps.Label1.Caption = txt3$
             '   Maps.Label2.Caption = txt4$
             '   End If
        ElseIf routeload = False And lResult > 0 Then
          worldmove = True
          Call goto_click
          worldmove = False
          End If
       End If

 If world = True And Maps.Timer2.Enabled = True And Dir(ramdrive + ":\mapwait.x") <> sEmpty Then
    'remove wait message after waiting a bit for blitting to finish
    timerwait = Timer + 0.1
    Do Until Timer > timerwait
       DoEvents
    Loop
    Kill ramdrive + ":\mapwait.x"
    End If

Exit Sub
positerror:
    If numtrys < 20 Then
       Resume
    Else
       Exit Sub
       End If
End Sub
Private Sub text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 And Maps.Timer2.Enabled = True Then Exit Sub
   If coordmode% <> 5 Then
      Maps.StatusBar1.Panels(2) = "X coordinate (change the coordinate system using RETURN key)"
   ElseIf (world = False And coordmode% = 5) Or (world = True And coordmode% = 2) Then
      Maps.StatusBar1.Panels(2) = "Distance from goto coordinates in kilometers"
      End If
End Sub
Private Sub text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 And Maps.Timer2.Enabled = True Then Exit Sub
   If coordmode% <> 5 Then
      Maps.StatusBar1.Panels(2) = "Y coordinate (change the coordinate system using RETURN key)"
   ElseIf (world = False And coordmode% = 5) Or (world = True And coordmode% = 2) Then
      Maps.StatusBar1.Panels(2) = "Azimut with respect to goto coordinates in degrees"
      End If
End Sub
Private Sub text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 And Maps.Timer2.Enabled = True Then Exit Sub
   If tblbuttons(1) = 0 Then
      Maps.StatusBar1.Panels(2) = "To activate the height option please place the DTM CD in the CD-ROM reader"
   ElseIf tblbuttons(1) = 1 Then
      Maps.StatusBar1.Panels(2) = "Height in meters"
      End If
End Sub
Private Sub text4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Text4.Visible = False Then
      Maps.StatusBar1.Panels(2) = sEmpty
   Else
      Maps.StatusBar1.Panels(2) = "Dip angle (degrees) with respect to the goto coordinates"
      End If
End Sub
Private Sub picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Maps.StatusBar1.Panels(2) = sEmpty
End Sub
Private Sub text6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Maps.StatusBar1.Panels(2) = "(Input) Y goto coordinate (change coordinate system using PGUP key)"
End Sub
Private Sub text5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Maps.StatusBar1.Panels(2) = "(Input) X goto coordinate (change coordinate system using PGUP key)"
End Sub
Private Sub text7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Maps.StatusBar1.Panels(2) = "Height in meters at goto coordinates (when DTM is activated)"
End Sub
Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Maps.StatusBar1.Panels(2) = sEmpty
End Sub
Private Sub statusbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If X >= StatusBar1.Panels(1).Width + StatusBar1.Panels(2).Width Then
      Maps.StatusBar1.Panels(2) = "Average of the remaining system and user resources"
      End If
End Sub
'Private Sub Picture4_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Label1_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Text1_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Label2_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Text2_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Label3_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Text3_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Label4_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Text4_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Label5_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Text5_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Label6_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Text6_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Label7_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub
'Private Sub Text7_MouseDown(Button As Integer, _
'   Shift As Integer, X As Single, Y As Single)
'   Select Case Button
'      Case 1 'left button
'         If Maps.Timer2.Enabled = True Then Exit Sub
'      Case Else
'   End Select
'End Sub

Private Sub Timer3_Timer()
   Dim lResource As Long
   'monitor the system resources
   lResource = FindWindow(vbNullString, "Resource Meter")
   If lResource > 0 Then
      resourcenum% = 0
      bRtn = EnumChildWindows(lResource, AddressOf EnumFunc2, 1) 'read captions
      If reboot = True Then
         Call MDIform_queryunload(i%, j%)
         End If
      End If
End Sub

Private Sub toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   X1 = 0: X2 = 0
   For i% = 1 To Toolbar1.Buttons.count
       X2 = X2 + Toolbar1.Buttons(i%).Width
       If X > X1 And X < X2 And Y > 0 And Y < Toolbar1.Height Then
         Maps.StatusBar1.Panels(2).Text = Toolbar1.Buttons(i%).ToolTipText
         If i% <= 5 Then
            exit3 = True
            End If
         If i% >= Toolbar1.Buttons.count - 4 Then exit3 = False
         Exit Sub
       End If
       X1 = X1 + Toolbar1.Buttons(i%).Width
   Next i%
   Maps.StatusBar1.Panels(2).Text = sEmpty 'default message
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Toolbar1_ButtonClick
' DateTime  : 3/28/2004 08:36
' Author    : Chaim Keller
' Purpose   : Main tool bar menus
'---------------------------------------------------------------------------------------
'
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
   Dim lResult As Long, lplacexist As Long, lResult2 As Long
   Dim xwin As Long, ywin As Long, winw As Long, winh As Long, winp As Long
   Dim nmsg As Long
   On Error GoTo Toolbar1_ButtonClick_Error '>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<

   Select Case Button.Key
     Case "DTMbut" 'enable/disenable height readings
        If tblbuttons(1) = 1 Then
           Toolbar1.Buttons(1).value = tbrPressed
           If world = False Then
              lResult = FindWindow(vbNullString, terranam$)
              If lResult > 0 Then 'remove topmost status
                 ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                 End If
              End If
           If world = True And mapPictureform.Visible = True Then
              ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
              End If
           response = MsgBox("This means that you won't be able to display heights!", _
                           vbExclamation + vbOKCancel, "Maps & More")
           If world = False And lResult > 0 Then 'restore topmost status
              ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
              End If
           If world = False And mapPictureform.Visible = True Then 'restore topmost status
              ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
              End If
           If world = True And mapPictureform.Visible = True Then
              ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
              End If
           If response = vbOK Then
              noheights = True
              CHMNEO = sEmpty
              Close
              mnuTrigDrag.Enabled = False
              mnuTrigUndo.Enabled = False
              If Not bAirPath Then
                 mnuCrossSection.Enabled = False
                 mnuFirstPoint.Enabled = False
                 mnuSecondPoint.Enabled = False
                 End If
              tblbuttons(1) = 0
              Toolbar1.Buttons(1).value = tbrUnpressed
              worldfil$ = sEmpty
              If world = True And mapEROSDTMwarn.Visible = True Then
                Unload mapEROSDTMwarn
                End If
           Else
'              skyDTMCDcheck.Value = vbChecked
'              Picture4.SetFocus
              Exit Sub
              End If
        Else 'means that want to use DTM CD to present heights,
             'check that its there
             On Error GoTo CDerror
             noheights = False
             CHMNEO = "XX"
d10:         If world = False Then
                myfile = Dir(israeldtm + ":\dtm\dtm-map.loc")
             Else
                myfile = Dir(worlddtm + ":\Gt30dem.gif")
                End If
             If myfile = sEmpty Then
                If world = False Then
                  lResult = FindWindow(vbNullString, terranam$)
                  If lResult > 0 Then
                     ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                     End If
                  End If
                  If world = True And mapPictureform.Visible = True Then
                      ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                      End If
                  If Not NoCDWarning Then
                     response = MsgBox("DTM CD not found, please load it into the CD drive.", _
                           vbCritical + vbOKCancel, "Maps & More")
                    NoCDWarning = True
                    End If
                If world = False And lResult > 0 Then
                     ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                     End If
                If world = True And mapPictureform.Visible = True Then
                      ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                      End If
                If response = vbOK Then
                   GoTo d10
                ElseIf response = vbCancel Then
                   tblbuttons(1) = 0
                   Toolbar1.Buttons(1).value = tbrUnpressed
                   noheights = True
                   End If
                End If
d15:       If noheights = False Then
              mnuTrigDrag.Enabled = True
              mnuTrigUndo.Enabled = True
              mnuCrossSection.Enabled = True
              mnuFirstPoint.Enabled = True
              mnuSecondPoint.Enabled = True
              If world = True Then
                 tblbuttons(1) = 1
                 Toolbar1.Buttons(1).value = tbrPressed
                 Exit Sub
                 End If
              filnum% = FreeFile
              Open israeldtm + ":\dtm\dtm-map.loc" For Input As #filnum%
              For i% = 1 To 3
                 Line Input #filnum%, doclin$
              Next i%
              N% = 0
              For i% = 4 To 54
                 Line Input #filnum%, doclin$
                 If i% Mod 2 = 0 Then
                    N% = N% + 1
                    For j% = 1 To 14
                       CHMAP(j%, N%) = Mid$(doclin$, 6 + (j% - 1) * 5, 2)
                    Next j%
                    End If
              Next i%
              Close #filnum%
              tblbuttons(1) = 1
              Toolbar1.Buttons(1).value = tbrPressed
              End If
           End If
'  Picture4.SetFocus
  Exit Sub
CDerror:
   If Err.Number = 71 Then
      If world = False Then
        lResult = FindWindow(vbNullString, terranam$)
        If lResult > 0 Then
           ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
           End If
        End If
      If world = True And mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
      response = MsgBox("Drive not ready, try again?", vbCritical + vbOKCancel, "Maps & More")
      If world = False And lResult > 0 Then
           ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
           End If
      If world = True And mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
      If response = vbOK Then
         Resume
      Else
         tblbuttons(1) = 0
         Toolbar1.Buttons(1).value = tbrUnpressed
         GoTo d15
         End If
    ElseIf Err.Number = 52 Or Err.Number = 53 Or Err.Number = 75 Or Err.Number = 76 Then
       If world = False Then
          lResult = FindWindow(vbNullString, terranam$)
          If lResult > 0 Then
             ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             End If
          End If
       If world = True And mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
       If Not NoCDWarning Then
          response = MsgBox("DTM CD not found, please load it into the CD drive.", _
                          vbCritical + vbOKCancel, "Maps & More")
          NoCDWarning = True
          End If
       If world = False And lResult > 0 Then
          ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
       If world = True And mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
          End If
       If response = vbOK Then
          Resume
       Else
          tblbuttons(1) = 0
          Toolbar1.Buttons(1).value = tbrUnpressed
          noheights = True
          GoTo d15
          End If
       End If
     Case "printbut"
        Dim bufxp(2, 4), bufyp(2, 4), bufwip(2, 4), bufhip(2, 4)
        If mapPictureform.Visible = False Then Exit Sub
        printing = True
        'set scales
        Printer.Width = mapPictureform.mapPicture.Width
        Printer.Height = mapPictureform.mapPicture.Height
        Printer.ScaleWidth = mapPictureform.mapPicture.ScaleWidth
        Printer.ScaleHeight = mapPictureform.mapPicture.ScaleHeight
        Printer.ScaleLeft = mapPictureform.mapPicture.ScaleLeft
        Printer.ScaleTop = mapPictureform.mapPicture.ScaleTop
        mapwit = mapwi
        mapwi2t = mapwi2
        maphit = maphi
        maphi2t = maphi2
        mapwi = sizex + 60
        mapwi2 = mapwi
        maphi = sizey + 60
        maphi2 = maphi

        obss% = 0
        If obstflag = True Then
           'ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
           obss% = 1
           obstflag = False
           printing = False
           End If
        Call blitpictures
        If obss% = 1 Then obstflag = True
        printing = True
       'blit to the printer buffer
        bufwip(1, 1) = bufwi(1, 1)
        bufwip(2, 1) = bufwi(2, 1)
        bufhip(1, 1) = bufhi(1, 1)
        bufhip(2, 1) = bufhi(2, 1)
        bufxp(1, 1) = bufx(1, 1) - printeroffset * mag
        bufxp(2, 1) = bufx(2, 1)
        bufyp(1, 1) = bufy(1, 1)
        bufyp(2, 1) = bufy(2, 1)

        bufhip(1, 2) = bufhi(1, 2)
        bufhip(2, 2) = bufhi(2, 2)
        bufwip(1, 2) = bufwi(1, 2)
        bufwip(2, 2) = bufwi(2, 2)
        bufxp(1, 2) = bufx(1, 2)
        bufxp(2, 2) = bufx(2, 2)
        bufyp(1, 2) = bufy(1, 2)
        bufyp(2, 2) = bufy(2, 2)

        bufxp(1, 3) = bufx(1, 3) - printeroffset * mag
        bufxp(2, 3) = bufx(2, 3)
        bufwip(1, 3) = bufwi(1, 3)
        bufwip(2, 3) = bufwi(2, 3)
        bufyp(1, 3) = bufy(1, 3) - printeroffset * mag
        bufyp(2, 3) = bufy(2, 3)
        bufhip(1, 3) = bufhi(1, 3)
        bufhip(2, 3) = bufhi(2, 3)

        bufxp(1, 4) = bufx(1, 4)
        bufxp(2, 4) = bufx(2, 4)
        bufwip(1, 4) = bufwi(1, 4)
        bufwip(2, 4) = bufwi(2, 4)
        bufyp(1, 4) = bufy(1, 4) - printeroffset * mag
        bufyp(2, 4) = bufy(2, 4)
        bufhip(1, 4) = bufhi(1, 4)
        bufhip(2, 4) = bufhi(2, 4)

        For i% = 1 To 4
         'check for nonsense widths,heights that cause program to bomb
          If bufwi(1, i%) <= 0 Then GoTo d100
          If bufhi(1, i%) <= 0 Then GoTo d100
          If bufwi(2, i%) <= 0 Then GoTo d100
          If bufhi(2, i%) <= 0 Then GoTo d100
          Printer.PaintPicture Maps.PictureClip1(bn%(i%)).Picture, bufxp(1, i%), bufyp(1, i%), bufwip(1, i%), bufhip(1, i%), bufxp(2, i%), bufyp(2, i%), bufwip(2, i%), bufhip(2, i%)
d100:   Next i%
        If tblbuttons(4) = 1 Then
           Call obstructions(Printer)
           End If
        If showroute = True Then
           Call showtheroute(Printer)
           End If

        mapwi = mapwit
        mapwi2 = mapwi2t
        maphi = maphit
        maphi2t = maphi2

'        obss% = 0
'        If obstflag = True Then
'           obss% = 1
'           obstflag = False
'           End If
        printing = False
        Call blitpictures
'        If obss% = 1 Then obstflag = True

         Printer.EndDoc
         printing = False
     Case "3Dexplorerbut" 'dump screen to ClipBoard
         'find if 3D explorer is enabled and full name of this window
         TdxhWnd = 0
         bRtn = EnumWindows(AddressOf EnumWndProc, lParam)
         If TdxhWnd = 0 Then
            If world = True And mapPictureform.Visible = True Then
               ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
               tblbuttons(3) = 1
               Maps.Toolbar1.Buttons(3).value = tbrPressed
               If ExplorerDir = sEmpty Then
                  If WinVer <> 5 And WinVer <> 261 Then
                     ExplorerDir = "c:\3dexplorer\"
                  ElseIf WinVer = 5 Or WinVer = 261 Then
                     ExplorerDir = "e:\3dexplorer\"
                     End If
                  End If
               If Dir(ExplorerDir & "3dxusa.exe") = sEmpty Then
                  response = InputBox("Can't find the 3D Explorer program" & vbLf & _
                                    "Enter the drive letter and directory." & vbLf & _
                                    "For example, d:\3dexplorer\ (include the last backslash).", _
                                    "3D Explorer directory", "d:\3DExplorer\", 6450)
                  If response = sEmpty Then
                     Exit Sub
                  Else
                     ExplorerDir = response
                     End If
                  End If
               rval = Shell(ExplorerDir & "3dxUSA.exe", vbNormalFocus)
               'wait until 3d explorer finishes loading
               waitime = Timer
               Do Until Timer > waitime + 2
                  DoEvents
               Loop
               'now find full name of this window
               TdxhWnd = 0
               bRtn = EnumWindows(AddressOf EnumWndProc, lParam)
               If TdxhWnd <> 0 Then
                  OverhWnd = FindWindow(vbNullString, "Overview")
                  'now move windows to proper position
                  xwin = -1
                  ywin = 99
                  winw = 390
                  winh = 409
                  winp = True
                  ret = MoveWindow(TdxhWnd, xwin, ywin, winw, winh, winp)
                  ret = MoveWindow(OverhWnd, xwin, ywin, winw, winh, winp)
                  ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                  ret = SetWindowPos(OverhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                  ret = SetWindowPos(TdxhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                  'now determine how far off the present coordinates of the 3D Explorer
                  'are off from the desired coordinates
                  'if they are off the present screen, then open Find window
                  iposit% = InStr(Tdxname, "-  ")
                  If iposit% <> 0 Then
                     lat3d = Val(Mid$(Tdxname, iposit% + 4, 2)) + Val(Mid$(Tdxname, iposit% + 8, 4)) / 60
                     lon3d = -(Val(Mid$(Tdxname, iposit% + 15, 3)) + Val(Mid$(Tdxname, iposit% + 19, 5)) / 60)
                     If Abs(Val(Maps.Text6.Text) - lat3d) > 0.37167 Or Abs(Val(Maps.Text5.Text) - lon3d > 0.37167) Then
                        'activate find window
                        Call BringWindowToTop(OverhWnd)
                        Call keybd_event(VK_F6, 0, 0, 0) 'activates alt key
                        Call keybd_event(VK_F6, 0, KEYEVENTF_KEYUP, 0)
                     Else
                        'move mouse cursor to right place and
                        'depress right mouse key to go there
                        dx1 = -1000 '-30 '30
                        dy1 = -1000 '-240 '60
                        Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
                        waitime = Timer + 0.001
                        Do Until Timer > waitime
                           DoEvents
                        Loop
                        dx1 = (Val(Maps.Text5.Text) - lon3d) * 516.6 + 96
                        dy1 = -(Val(Maps.Text6.Text) - lat3d) * 516.6 + 156
                        Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
                        waitime = Timer + 0.001
                        Do Until Timer > waitime
                           DoEvents
                        Loop
                        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) 'move mouse to Location item
                        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0) 'move mouse to Location item
                        End If
                     End If
               Else
                  ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
                  End If
               End If
         ElseIf world = True And TdxhWnd <> 0 Then
            'bring 3d explorer to top of z order
            ret = SetWindowPos(TdxhWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
            Exit Sub
         Else
            Exit Sub
            End If
     Case "obstructbut"
          On Error GoTo obserrhandler
          'If world = True Then Exit Sub
          If tblbuttons(4) = 0 Then
             tblbuttons(4) = 1
             Maps.Toolbar1.Buttons(4).value = tbrPressed
             'set filters
             ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
             CommonDialog1.Filter = "Sunrise/Sunset profile files (*.p*)|*.p*|" + _
                                 "All files (*.*)|*.*|"
             'specify the default flter
             CommonDialog1.FilterIndex = 7
'             lResult = FindWindow(vbNullString, "Open")
             'display the open dialog box
             CommonDialog1.FileName = drivcities$ + "*.p*"
             CommonDialog1.ShowOpen
             'read the files coordinates, goto there, and then plot the obstructions
             Screen.MousePointer = vbHourglass
             obsfile$ = CommonDialog1.FileName
             Maps.Caption = Maps.Caption + "  (obstruction file: " + obsfile$ + ")"
             obsfilnum% = FreeFile
             ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
             myfile = Dir(obsfile$)
             If myfile <> sEmpty Then
                Open obsfile$ For Input As obsfilnum%
                Line Input #obsfilnum%, doclin$
                Input #obsfilnum%, kmxob, kmyob, hgtob, Aob, Bob, Cob, Dob, Eob
                'Close #obsfilnum%
                If world = False Then
                    coordmode2% = 1
                    Maps.Label5.Caption = "ITMx"
                    Maps.Label6.Caption = "ITMy"
                    Maps.Text7.Text = hgtob - 1.8
                    Maps.Text5.Text = kmxob * 1000
                    Maps.Text6.Text = kmyob * 1000 + 1000000
                    kmxc = Maps.Text5.Text: kmyc = Maps.Text6.Text
                    kmxobs = kmxc: kmyobs = kmyc
                    kmxsky = kmxc: kmysky = kmyc
                Else
                   coordmode2% = 2
                   Maps.Label5.Caption = "long."
                   Maps.Label6.Caption = "latit."
                   Maps.Text7.Text = hgtob - 1.8
                   Maps.Text5.Text = -kmyob
                   Maps.Text6.Text = kmxob
                   lon = Maps.Text5.Text
                   lat = Maps.Text6.Text
                   lonobs = lon: latobs = lat
                   lonc = lon + (deglog / sizewx) * (mapwi - mapwi2 + mapxdif) / 2 + fudx
                   latc = lat - (deglat / sizewy) * (maphi - maphi2 + mapydif) / 2 + fudy
                   End If
                obstflag = True
                'now load rest of file into arrays
                obsnum% = 0
                Do Until EOF(obsfilnum%)
                   Input #obsfilnum%, aziob, vaob, kmxob, kmyob, c, D
                   If kmxob = 0 Or kmyob = 0 Then
                      MsgBox "This appears to be surveyor results!", _
                             vbExclamation + vbOKOnly, "Maps&More"
                      Close #obsfilnum%
                      Exit Sub
                      End If
                   obsnum% = obsnum% + 1
                   'If obsnum% > travelmax% Then
                   '   response = MsgBox("This obstruction file has too many points!  Sorry, it can't be viewed (unless you change the value of travelmax%).", vbOKOnly + vbExclamation, "Maps & More")
                   '   Maps.Toolbar1.Buttons(4) = tbrUnpressed
                   '   tblbuttons(4) = 0
                   '   obstflag = False
                   '   Exit Sub
                   '   End If
                   ReDim Preserve obs(2, obsnum%)
                   If world = False Then
                      obs(1, obsnum%) = kmxob * 1000
                      obs(2, obsnum%) = kmyob * 1000 + 1000000
                   Else
                      obs(1, obsnum%) = kmxob
                      obs(2, obsnum%) = kmyob 'negative for East longitude
                      End If
                Loop
                Close #obsfilnum%
                Call goto_click
                Screen.MousePointer = vbDefault
                End If
          Else
             'user decided to undepress the obstruction button
             Maps.Toolbar1.Buttons(4).value = tbrUnpressed
             tblbuttons(4) = 0
             obstflag = False
             'Close #obsfilnum%
             If mapPictureform.Visible = True Then
                Call blitpictures 'erase the obstruction lines
                End If
             Maps.Caption = mapcapold$
             End If
          Exit Sub
obserrhandler:
        'user pressed cancel button
         Maps.Toolbar1.Buttons(4).value = tbrUnpressed
         tblbuttons(4) = 0
         obstflag = False
         Exit Sub
     Case "placebut" 'popup stored places
        If mapPLACfm.Visible = False Then
           tblbuttons%(9) = 1
           Toolbar1.Buttons(9).value = tbrPressed
           Screen.MousePointer = vbHourglass
        Else
           tblbuttons%(9) = 0
           Toolbar1.Buttons(9).value = tbrUnpressed
           mapPLACfm.Visible = False
           Unload mapPLACfm
           Exit Sub
           End If
        If world = False Then
           lResult = FindWindow(vbNullString, terranam$)
           If lResult > 0 Then
              ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
              End If
           End If
        If lplac% = 0 Then
           If tblbuttons%(4) = 1 Then 'then obstruction button depreesed
              response = MsgBox("You must depress the OBSTRUCTION button before accessing the place list.", vbInformation + vbOKCancel, "Maps & More")
              If response = vbOK Then
                 Do Until tblbuttons%(4) = 0
                    DoEvents
                 Loop
              Else
                 Exit Sub
                 End If
              End If
           mapPLACfm.Visible = True
           ret = SetWindowPos(mapPLACfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
           Screen.MousePointer = vbDefault
           lplac% = 1
        ElseIf lplac% = 1 Then 'bring places to top of z order
           lplacexist = FindWindow(vbNullString, mapPLACfm.Caption)
           ret = BringWindowToTop(lplacexist)
           End If
     Case "timerbut" 'change timer setting
        If mapCHANGETIMEfm.Visible = False Then
           mapCHANGETIMEfm.Visible = True
           Toolbar1.Buttons(21).value = tbrPressed
           tblbuttons%(21) = 1
        Else
           mapCHANGETIMEfm.Visible = False
           Unload mapCHANGETIMEfm
           Toolbar1.Buttons(21).value = tbrUnpressed
           tblbuttons%(21) = 0
           End If
     Case "gotobut" 'goto inputed coordinates
       Screen.MousePointer = vbHourglass
       gotobutton = True
       Call goto_click
       gotobutton = False
       Screen.MousePointer = vbDefault
     Case "map50but" 'display 1:50000 topo maps of Eretz Israel
       Call map50butsub
     Case "map400but" 'display 1:400000 relief maps of Eretz Israel
       topofm.Enabled = False
       If mapPictureform.Visible = False Or map50 = True Or world = True Then
          If map50 = True Then
             map50 = False
             Toolbar1.Buttons(7).value = tbrUnpressed
             tblbuttons(7) = 0
             End If
          If world = True Then
             world = False
             mapPictureform.Visible = False
             tblbuttons(3) = 0
             tblbuttons(8) = 0
             Maps.Toolbar1.Buttons(3).value = tbrUnpressed
             Maps.Toolbar1.Buttons(3).Enabled = False
             Toolbar1.Buttons(8).value = tbrUnpressed
             mapPictureform.Width = sizex + 60 '60 is the size (pixels) of the borders
             mapPictureform.Height = sizey + 60
             mapPictureform.mapPicture.Width = sizex
             mapPictureform.mapPicture.Height = sizey
             mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
             mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
             If world = True Then
                mapxdif = mapxdif + 35
                mapydif = mapydif + 35
                End If
             mapwi = mapPictureform.Width
             maphi = mapPictureform.Height
             If mapEROSDTMwarn.Visible = True Then
               Unload mapEROSDTMwarn
               End If
             End If
          If noheights = False Then
            mnuCrossSection.Enabled = True
            mnuFirstPoint.Enabled = True
            mnuSecondPoint.Enabled = True
            End If
          map400 = True
          For i% = 2 To 15
              Toolbar1.Buttons(i%).Enabled = True
          Next i%
          Toolbar1.Buttons(26).value = tbrUnpressed
          Toolbar1.Buttons(27).value = tbrUnpressed
          If RdHalYes Then 'enable sunrise/sunset calculations
            Toolbar1.Buttons(26).Enabled = True
            Toolbar1.Buttons(27).Enabled = True
          Else
            Toolbar1.Buttons(26).Enabled = False
            Toolbar1.Buttons(27).Enabled = False
            End If

          Toolbar1.Buttons(20).Enabled = True
          Toolbar1.Buttons(21).Enabled = True
          searchfm.Enabled = True
          Combo1.Enabled = True
          Picture4.Visible = True
          coordmode% = 1
          Maps.Text4.Visible = False
          Maps.Label4.Visible = False
          openbatfm.Enabled = True
'''          coordmode2% = 1
          Label1.Caption = "ITMx"
          Label2.Caption = "ITMy"
          Label5.Caption = "ITMx"
          Label6.Caption = "ITMy"
          Maps.Text1.Text = "0"
          Maps.Text2.Text = "0"
          Maps.Text3.Text = "0"
          Maps.Text5.Text = kmxc
          Maps.Text6.Text = kmyc
          Maps.Text7.Text = hgt400c
          If Maps.Text7.Text = sEmpty Then
             hgt50c = 0: hgtpos = 0
             Maps.Text7.Text = "0"
             End If
          lResult = FindWindow(vbNullString, terranam$)
          If lResult > 0 And terranam$ <> sEmpty Then
             For i% = 18 To 21
                Toolbar1.Buttons(i%).Enabled = True
             Next i%
             Loadfm.Enabled = True
             If Dir(ramdrive + ":\travlog.x") <> sEmpty Then recoverroutefm.Enabled = True
             End If
          appendfrm.Enabled = True
          Toolbar1.Buttons(6).value = tbrPressed
          tblbuttons(6) = 1
          Toolbar1.Buttons(17).Enabled = True
          Call loadpictures  'load appropriate map tiles into off-screen buffers
          Call blitpictures   'blit desired portions of the off-screen buffers to the screen
          If kmx50c = 0 And kmy50c = 0 And kmx400c = 0 And kmy400c = 0 Then
             kmx50c = kmxc: kmy50c = kmyc: hgt50c = hgt
             kmx400c = kmxc: kmy400c = kmyc: hgt400c = hgt
             'convert to sky coordinates
             'mode% = 1
             'kmxo = kmxc: kmyo = kmyc
             kmxsky = kmxc: kmysky = kmyc
             'Call ITMSKY(kmxo, kmyo, T1, T2, mode%)
             'Maps.Text5.Text = T1
             'Maps.Text6.Text = T2
             'Maps.Text7.Text = hgtpos
             End If
          If mapPictureform.Visible = True Then
             ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
             End If
       Else
          map400 = False
          If Not bAirPath Then
            mnuCrossSection.Enabled = False
            mnuFirstPoint.Enabled = False
            mnuSecondPoint.Enabled = False
            End If
          openbatfm.Enabled = False
          Toolbar1.Buttons(6).value = tbrUnpressed
          tblbuttons(6) = 0
          Loadfm.Enabled = False
          recoverroutefm.Enabled = False
          mapPictureform.Visible = False
          For i% = 9 To 15
             Toolbar1.Buttons(i%).Enabled = False
             tblbuttons(i%) = 0
          Next i%
          For i% = 18 To 27
             Toolbar1.Buttons(i%).Enabled = False
             tblbuttons(i%) = 0
          Next i%
          appendfrm.Enabled = False
          Toolbar1.Buttons(2).Enabled = False
          Toolbar1.Buttons(4).Enabled = False
          Toolbar1.Buttons(20).value = tbrUnpressed
          tblbuttons(20) = 0
          Combo1.Enabled = False
          Picture4.Visible = False
          searchfm.Enabled = False
          If tblbuttons(4) = 1 Then
             Maps.Toolbar1.Buttons(4).value = tbrUnpressed
             tblbuttons(4) = 0
             obstflag = False
             End If
          End If
     Case "worldmap"
       topofm.Enabled = False
       fudx = 0
       fudy = 0
       Toolbar1.Buttons(8).value = tbrPressed
       tblbuttons(8) = 1
       If mapPictureform.Visible = False Or map50 = True Or map400 = True Then
          pix600fm.Enabled = False
          pix1200fm.Enabled = False
          topofm.Enabled = False
          importfm.Enabled = True
          importmapfm.Enabled = True
          resetoriginfm.Enabled = True
          If noheights = False Then
            mnuCrossSection.Enabled = True
            mnuFirstPoint.Enabled = True
            mnuSecondPoint.Enabled = True
            End If
          mapimport = False
          For i% = 1 To 9
             picold$(i%) = sEmpty
          Next i%

          'If tblbuttons(4) = 1 Then
          '   obstflag = False
          '   End If
          If map50 = True Then
             map50 = False
             End If
          If map400 = True Then
             map400 = False
             End If
          Picture4.Visible = True
          Combo1.Enabled = True
          coordmode% = 2
          Maps.Text4.Visible = False
          Maps.Label4.Visible = False
          Label1.Caption = "long."
          Label2.Caption = "latit."
          Label5.Caption = "long."
          Label6.Caption = "latit."
          If lon = 0 And lat = 0 Then
             lon = 35.2385
             lat = 31.805042
             hgtworld = 762
             End If
          openbatfm.Enabled = True
          Maps.Text5.Text = Format(lon, "###0.0#####")
          Maps.Text6.Text = Format(lat, "##0.0#####")
          Maps.Text7.Text = hgtworld
          lono = lon
          lato = lat
          lg = lon
          lt = lat
          'lgdeg = Fix(lg)
          'lgmin = Abs(Fix((lg - Fix(lg)) * 60))
          'lgsec = Abs(((lg - Fix(lg)) * 60 - Fix((lg - Fix(lg)) * 60)) * 60)
          'ltdeg = Fix(lt)
          'ltmin = Abs(Fix((lt - Fix(lt)) * 60))
          'ltsec = Abs(((lt - Fix(lt)) * 60 - Fix((lt - Fix(lt)) * 60)) * 60)
          'If ltdeg = 0 And lt < 0 Then
          '   Maps.Text6.Text = "-" + Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
          'Else
          '   Maps.Text6.Text = Str$(ltdeg) + "°" + Str$(ltmin) + "'" + Mid$(Str$(ltsec), 1, 6) + """"
          '   End If
          'If lgdeg = 0 And lg < 0 Then
          '   Maps.Text5.Text = "-" + Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
          'Else
          '   Maps.Text5.Text = Str$(lgdeg) + "°" + Str$(lgmin) + "'" + Mid$(Str$(lgsec), 1, 6) + """"
          '   End If
          Loadfm.Enabled = False
          recoverroutefm.Enabled = False
          Maps.Text1.Text = "0"
          Maps.Text2.Text = "0"
          Maps.Text3.Text = "0"
          coordmode2% = 2
          lResult = FindWindow(vbNullString, terranam$)
          If lResult > 0 And terranam$ <> sEmpty Then 'stop terraviewer
          'close the terraviewer
             ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
             Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
             Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
             Call keybd_event(Asc("F"), 0, 0, 0)  'goes into Settings menu
             Call keybd_event(Asc("F"), 0, KEYEVENTF_KEYUP, 0)
             Call keybd_event(Asc("E"), 0, 0, 0)  'goes into Settings menu
             Call keybd_event(Asc("E"), 0, KEYEVENTF_KEYUP, 0)
             Toolbar1.Buttons(17).value = tbrUnpressed
             For i% = 17 To 19
                Toolbar1.Buttons(i%).Enabled = False
             Next i%
          Else
             End If
          mapPictureform.Visible = False
          world = True
          pixwi = 594 'size of Eretz Israel bitmaps in pixels
          pixhi = 594
          pixwwi = 604 '603 '599 603 '604
          pixwhi = 604 '602 '598 602 '604
          sizex = Screen.TwipsPerPixelX * pixwi '# twips in half of picture=8850/2
          sizey = Screen.TwipsPerPixelY * pixhi '=8850/2
          sizewx = Screen.TwipsPerPixelX * pixwwi '# twips in half of picture=8850/2
          sizewy = Screen.TwipsPerPixelY * pixwhi '=8850/2
          mapPictureform.Width = sizewx + 60 '60 is the size (pixels) of the borders
          mapPictureform.Height = sizewy + 60
          mapPictureform.mapPicture.Width = sizewx
          mapPictureform.mapPicture.Height = sizewy
          mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
          mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
          If world = True Then
             mapxdif = mapxdif + 35
             mapydif = mapydif + 35
             End If
          mapwi = mapPictureform.Width
          maphi = mapPictureform.Height
          mapwi = mapPictureform.Width
          maphi = mapPictureform.Height
          Call loadpictures  'load appropriate map tiles into off-screen buffers
          Call blitpictures   'blit desired portions of the off-screen buffers to the screen
          Toolbar1.Buttons(17).Enabled = False
          Toolbar1.Buttons(2).Enabled = False
          Toolbar1.Buttons(4).Enabled = True 'False
          Toolbar1.Buttons(9).Enabled = True
          Toolbar1.Buttons(10).Enabled = True
          For i% = 12 To 15
             Toolbar1.Buttons(i%).Enabled = True
          Next i%
          For i% = 18 To 21
             Toolbar1.Buttons(i%).Enabled = True
             Toolbar1.Buttons(i%).value = tbrUnpressed
             tblbuttons(i%) = 0
          Next i%
          Toolbar1.Buttons(26).Enabled = True
          Toolbar1.Buttons(27).Enabled = True
          tblbuttons(20) = 0
          Maps.Toolbar1.Buttons(4).value = tbrUnpressed
          tblbuttons(4) = 0
          Toolbar1.Buttons(7).value = tbrUnpressed
          tblbuttons(7) = 0
          Toolbar1.Buttons(6).value = tbrUnpressed
          tblbuttons(6) = 0
          Toolbar1.Buttons(23).Enabled = False
          Toolbar1.Buttons(23).value = tbrUnpressed
          tblbuttons(23) = 0
          Toolbar1.Buttons(24).Enabled = False
          Toolbar1.Buttons(24).value = tbrUnpressed
          tblbuttons(24) = 0
          Toolbar1.Buttons(25).value = tbrUnpressed
          Toolbar1.Buttons(25).Enabled = False
          tblbuttons(25) = 0
          showroute = False
          If mapPictureform.Visible = True Then
             ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
             End If
          Timer2.Interval = 7500
          Loadfm.Enabled = True
          If Dir(ramdrive + ":\travlog.x") <> sEmpty Then recoverroutefm.Enabled = True
          mihr50fm.Caption = "&1.   500 mi/hr"
          mihr60fm.Caption = "&2.   600 mi/hr"
          mihr70fm.Caption = "&3.   700 mi/hr"
          mihr80fm.Caption = "&4.   800 mi/hr"
          mihr90fm.Caption = "&5.   900 mi/hr"
          mihr100fm.Caption = "&6.  1000 mi/hr"
          mihr110fm.Caption = "&7.  1100 mi/hr"
          mihr120fm.Caption = "&8.  1200 mi/hr"
          searchfm.Enabled = True
          Maps.Toolbar1.Buttons(3).Enabled = True
       Else
          world = False
          openbatfm.Enabled = False
          pix600fm.Enabled = True
          pix1200fm.Enabled = True
          topofm.Enabled = True
          importfm.Enabled = False
          importmapfm.Enabled = False
          importcenterfm.Enabled = False
          resetoriginfm.Enabled = False
          If Not bAirPath Then
            mnuCrossSection.Enabled = False
            mnuFirstPoint.Enabled = False
            mnuSecondPoint.Enabled = False
            End If
          mapimport = False
          For i% = 1 To 9
             picold$(i%) = sEmpty
          Next i%
          If mapEROSDTMwarn.Visible = True Then
             Unload mapEROSDTMwarn
             End If
          mapPictureform.Visible = False
          mapPictureform.Width = sizex + 60 '60 is the size (pixels) of the borders
          mapPictureform.Height = sizey + 60
          mapPictureform.mapPicture.Width = sizex
          mapPictureform.mapPicture.Height = sizey
          mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
          mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
          If world = True Then
             mapxdif = mapxdif + 35
             mapydif = mapydif + 35
             End If
          mapwi = mapPictureform.Width
          maphi = mapPictureform.Height
          Toolbar1.Buttons(8).value = tbrUnpressed
          tblbuttons(8) = 0
          Toolbar1.Buttons(17).Enabled = True
          mapPictureform.Visible = False
          For i% = 9 To 15
             Toolbar1.Buttons(i%).Enabled = False
             tblbuttons(i%) = 0
          Next i%
          For i% = 18 To 21
             Toolbar1.Buttons(i%).Enabled = False
             Toolbar1.Buttons(i%).value = tbrUnpressed
             tblbuttons(i%) = 0
          Next i%
          Toolbar1.Buttons(20).Enabled = False
          showroute = False
          travelmode = False
          Toolbar1.Buttons(2).Enabled = False
          If tblbuttons(4) = 1 Then
             Maps.Toolbar1.Buttons(4).value = tbrUnpressed
             tblbuttons(4) = 0
             obstflag = False
             End If
          Toolbar1.Buttons(26).Enabled = False
          Toolbar1.Buttons(27).Enabled = False
          Toolbar1.Buttons(26).value = tbrUnpressed
          Toolbar1.Buttons(27).value = tbrUnpressed
          Combo1.Enabled = False
          Picture4.Visible = False
          Loadfm.Enabled = False
          recoverroutefm.Enabled = False
          mihr50fm.Caption = "&1.    50 mi/hr"
          mihr60fm.Caption = "&2.    60 mi/hr"
          mihr70fm.Caption = "&3.    70 mi/hr"
          mihr80fm.Caption = "&4.    80 mi/hr"
          mihr90fm.Caption = "&5.    90 mi/hr"
          mihr100fm.Caption = "&6.   100 mi/hr"
          mihr110fm.Caption = "&7.   110 mi/hr"
          mihr120fm.Caption = "&8.   120 mi/hr"
          searchfm.Enabled = False
          tblbuttons(3) = 0
          Maps.Toolbar1.Buttons(3).value = tbrUnpressed
          Maps.Toolbar1.Buttons(3).Enabled = False
          End If
     Case "rightbut"
       If mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          End If
       If mapPictureform.Visible = False Then Exit Sub
       If tblbuttons%(13) = 0 Then
          If tblbuttons%(12) = 1 Then
             Toolbar1.Buttons(12).value = tbrUnpressed
             tblbuttons%(12) = 0
             End If
          Toolbar1.Buttons(13).value = tbrPressed
          tblbuttons%(13) = 1
          Timer1.Enabled = True
          Exit Sub
          End If
       If tblbuttons%(13) = 1 Then
          tblbuttons%(13) = 0
          Toolbar1.Buttons(13).value = tbrUnpressed
          For i% = 12 To 15
             If tblbuttons%(i%) <> 0 Then Exit Sub
          Next i%
          Timer1.Enabled = False
          End If
     Case "leftbut"
       If mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          End If
       If mapPictureform.Visible = False Then Exit Sub
       If tblbuttons%(12) = 0 Then
          If tblbuttons%(13) = 1 Then
             Toolbar1.Buttons(13).value = tbrUnpressed
             tblbuttons%(13) = 0
             End If
          Toolbar1.Buttons(12).value = tbrPressed
          tblbuttons%(12) = 1
          Timer1.Enabled = True
          Exit Sub
          End If
       If tblbuttons%(12) = 1 Then
          tblbuttons%(12) = 0
          Toolbar1.Buttons(12).value = tbrUnpressed
          For i% = 12 To 15
             If tblbuttons%(i%) <> 0 Then Exit Sub
          Next i%
          Timer1.Enabled = False
          End If
     Case "downbut"
       If mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          End If
       If mapPictureform.Visible = False Then Exit Sub
       If tblbuttons%(14) = 0 Then
          If tblbuttons%(15) = 1 Then
             Toolbar1.Buttons(15).value = tbrUnpressed
             tblbuttons%(15) = 0
             End If
          Toolbar1.Buttons(14).value = tbrPressed
          tblbuttons%(14) = 1
          Timer1.Enabled = True
          Exit Sub
          End If
       If tblbuttons%(14) = 1 Then
          tblbuttons%(14) = 0
          Toolbar1.Buttons(14).value = tbrUnpressed
          For i% = 12 To 15
             If tblbuttons%(i%) <> 0 Then Exit Sub
          Next i%
          Timer1.Enabled = False
          End If
     Case "upbut"
       If mapPictureform.Visible = True Then
          ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          End If
       If mapPictureform.Visible = False Then Exit Sub
       If tblbuttons%(15) = 0 Then
          If tblbuttons%(14) = 1 Then
             Toolbar1.Buttons(14).value = tbrUnpressed
             tblbuttons%(14) = 0
             End If
          Toolbar1.Buttons(15).value = tbrPressed
          tblbuttons%(15) = 1
          Timer1.Enabled = True
          Exit Sub
          End If
       If tblbuttons%(15) = 1 Then
          tblbuttons%(15) = 0
          Toolbar1.Buttons(15).value = tbrUnpressed
          For i% = 12 To 15
             If tblbuttons%(i%) <> 0 Then Exit Sub
          Next i%
          Timer1.Enabled = False
          End If
     Case "terrabut"
       If terranam$ = sEmpty Then terranam$ = "TerraExplorer - " + terradir$ + "\Israel9.teh"
       'If terranam$ = sEmpty Then terranam$ = "TerraViewer Basic - e:\terraviewer\default.hdr"
       lResult = FindWindow(vbNullString, terranam$)
       If lResult = 0 Then
          lResult2 = FindWindow(vbNullString, "TerraExplorer - ")
          If lResult2 > 0 Then terranam$ = "TerraExplorer - "
          End If
       If lResult > 0 Or lResult2 > 0 Then
          For i% = 18 To 19
             Toolbar1.Buttons(i%).Enabled = False
             Toolbar1.Buttons(i%).value = tbrUnpressed
             tblbuttons(i%) = 0
          Next i%
          Toolbar1.Buttons(23).value = tbrUnpressed
          Toolbar1.Buttons(23).Enabled = False
          tblbuttons(23) = 0
          Toolbar1.Buttons(24).value = tbrUnpressed
          Toolbar1.Buttons(24).Enabled = False
          tblbuttons(24) = 0
          Toolbar1.Buttons(25).value = tbrUnpressed
          Toolbar1.Buttons(25).Enabled = False
          tblbuttons(25) = 0
          Toolbar1.Buttons(17).value = tbrUnpressed
          tblbuttons(17) = 0
          Toolbar1.Refresh
          showroute = False
          Maps.Timer2.Enabled = False
          skyleftjump = False
          skymove = False
          'close the terraviewer
          ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
          Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
          Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
          'waitime = Timer
          'Do Until Timer > waitime + 0.01
          '   DoEvents
          'Loop
          Call keybd_event(Asc("F"), 0, 0, 0)  'goes into Settings menu
          Call keybd_event(Asc("F"), 0, KEYEVENTF_KEYUP, 0)
          Call keybd_event(Asc("X"), 0, 0, 0)  'goes into Settings menu
          Call keybd_event(Asc("X"), 0, KEYEVENTF_KEYUP, 0)
          Loadfm.Enabled = False
          recoverroutefm.Enabled = False
          'lResult = FindWindow(vbNullString, terranam$)
          ' ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
           ''Call keybd_event(KF_ALTDOWN, 0, 0, 0)
           'timwait = Timer
           'Do Until Timer > timwait + 0.1
           '  DoEvents
           'Loop
           'Call keybd_event(VK_SNAPSHOT, 0, 0, 0)
       Else
          If mapPictureform.Visible = True Then
             For i% = 18 To 21
                Toolbar1.Buttons(i%).Enabled = True
             Next i%
             Loadfm.Enabled = True
             If Dir(ramdrive + ":\travlog.x") <> sEmpty Then recoverroutefm.Enabled = True
             End If
          Toolbar1.Buttons(17).value = tbrPressed
          tblbuttons(17) = 1
          ChDrive Mid$(terradir$, 1, 1)
          ChDir terradir$  'd:\terraviewer\"
          taskID = Shell(terradir$ + "\TerraExplorer.exe --" + terradir$ + "\Israel9.teh", vbNormalFocus)
          'taskID = Shell("e:\terraviewer\TerraExplorer.exe", vbNormalFocus)
'           taskID = Shell("d:\terraviewer\terraviewer.exe --e:\terraviewer\default.hdr", vbNormalFocus)
           'now move and shape the window
          'waitime = Timer
          'Do Until Timer > waitime + 1
          '   DoEvents
          'Loop
          waitime = Timer
          lResult = FindWindow(vbNullString, "TerraExplorer - " + terradir$ + "\Israel9.teh")
          Do Until lResult <> 0 Or Timer > waitime + 120
             lResult = FindWindow(vbNullString, "TerraExplorer - " + terradir$ + "\Israel9.teh")
             DoEvents
          Loop

          terranam$ = "TerraExplorer - " + terradir$ + "\Israel9.teh"
'          lResult = FindWindow(vbNullString, "TerraViewer Basic - e:\terraviewer\default.hdr")
'         terranam$ = "TerraViewer Basic - e:\terraviewer\default.hdr"
          If lResult = 0 Then
             lResult = FindWindow(vbNullString, "TerraExplorer - ")
             'lResult = FindWindow(vbNullString, "TerraExplorer - e:\terraviewer\Israel9.teh")
'             lResult = FindWindow(vbNullString, "TerraViewer Basic - .\default.hdr")
             If lResult > 0 Then
                terranam$ = "TerraExplorer - "
'                terranam$ = "TerraExplorer - e:\terraviewer\Israel9.teh"
'                terranam$ = "TerraViewer Basic - .\default.hdr"
                End If
             End If
          xwin = 280
          ywin = 73 + 26 '0
          winw = 520 '515 '510
          winh = 472 '475 '472
          winp = True
          ret = MoveWindow(lResult, xwin, ywin, winw, winh, winp)

          ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          waitime = Timer 'now get rid of navigation map
          Do Until Timer > waitime + 0.1
             DoEvents
          Loop
          'hide navigation map as default
          'Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
          'Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
          'Call keybd_event(Asc("F"), 0, 0, 0)  'goes into view menu

          'Call keybd_event(Asc("F"), 0, KEYEVENTF_KEYUP, 0)
          'Call keybd_event(49, 0, 0, 0)  'opens file
          'Call keybd_event(49, 0, KEYEVENTF_KEYUP, 0)
          'terranam$ = "TerraExplorer - e:\terraviewer\Israel9.teh"
          Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
          Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
          Call keybd_event(Asc("V"), 0, 0, 0)  'goes into view menu
          Call keybd_event(Asc("V"), 0, KEYEVENTF_KEYUP, 0)
          Call keybd_event(Asc("N"), 0, 0, 0)    'calls for the navigation map
          Call keybd_event(Asc("N"), 0, KEYEVENTF_KEYUP, 0)
          Call keybd_event(VK_RETURN, 0, 0, 0)
          Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)

          'set high frame rate as default (already set in new version)
          'Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
          'Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
          'Call keybd_event(Asc("S"), 0, 0, 0)  'goes into Settings menu
          'Call keybd_event(Asc("S"), 0, KEYEVENTF_KEYUP, 0)
          'Call keybd_event(VK_RETURN, 0, 0, 0)
          'Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
          'Call keybd_event(VK_RETURN, 0, 0, 0)
          'Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)

          'make the TerraViewer toolbar disappear
          Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
          Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
          Call keybd_event(Asc("V"), 0, 0, 0)  'goes into Settings menu
          Call keybd_event(Asc("V"), 0, KEYEVENTF_KEYUP, 0)
          'in old version needed to press "T, in newer version press M"
          Call keybd_event(Asc("M"), 0, 0, 0)  'goes into Settings menu
          Call keybd_event(Asc("M"), 0, KEYEVENTF_KEYUP, 0)

          'Call keybd_event(Asc("T"), 0, 0, 0)  'goes into Settings menu
          'Call keybd_event(Asc("T"), 0, KEYEVENTF_KEYUP, 0)
    '      bRtn = EnumChildWindows(lResult, AddressOf EnumFunc, lParam) 'read captions

          'goto SKY coordinates that were previously inputed
           If Maps.Label5.Caption <> "long." And Maps.Text5.Text <> sEmpty And Picture4.Visible = True Then
sky200:      Call skyTERRAgoto
             'wait to get there, then go to minimum elevation
             '****old settings'
             'waitime = Timer
             'Do Until Timer > waitime + 1
             '   DoEvents
             'Loop
             ''first move pointer to TerraViewer window in order to activate it
             'dx1 = 30 '-30
             'dy1 = 240 ' -60
             'Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item
             'For i% = 1 To 30 'hold down the keys for a bit
             '   Call keybd_event(VK_SHIFT, 0, 0, 0)
             '   Call keybd_event(Asc("X"), 0, 0, 0)  'goes into Settings menu
             'Next i%
             'Call keybd_event(Asc("X"), 0, KEYEVENTF_KEYUP, 0)
             'Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
             'move pointer back to original position
             dx1 = 70 '-30 '30
             dy1 = -211 '-240 '60
             If WinVer = 261 Then
                dx1 = 35
                dy1 = -115
                End If
             Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
             End If
         End If
     Case "Animatbut"
         If routeload = True Then Exit Sub
         If world = True Then
            If tblbuttons(18) = 0 Then
               tblbuttons(18) = 1
               Toolbar1.Buttons(18).value = tbrPressed
               Maps.Timer2.Enabled = True
            Else
               Maps.Timer2.Enabled = False
               tblbuttons(18) = 0
               Toolbar1.Buttons(18).value = tbrUnpressed
               End If
         Else
          If tblbuttons(18) = 0 Then
             tblbuttons(18) = 1
             Toolbar1.Buttons(18).value = tbrPressed
             If Toolbar1.Buttons(17).Enabled = True Then
                Toolbar1.Buttons(17).Enabled = False
                End If
             'lResult = FindWindow(vbNullString, terranam$)
             'ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             'ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)

             'response = MsgBox("If you wan't the maps to follow the TerraExplorer then keep " + _
             '         "the mouse cursor over the TerraExplorer X: or Y: boxes.  If you " + _
             '         "to move the cursor to another place, it is advisable to depress " + _
             '         "the Animate button.", vbOKCancel + vbInformation, "Maps & More")
             'ret = SetWindowPos(mapPictureform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
             'ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)

             'If response = vbCancel Then Exit Sub

             'give the user 3 seconds to get the cursor to the TerraExplorer X:,Y: edit boxes
             timerwait = Timer + 3
             Do Until Timer > timerwait
               DoEvents
             Loop

             Maps.Timer2.Enabled = True
          Else
             Maps.Timer2.Enabled = False
             Toolbar1.Buttons(17).Enabled = True
             tblbuttons(18) = 0
             Toolbar1.Buttons(18).value = tbrUnpressed
             Screen.MouseIcon = LoadPicture("")
             Screen.MousePointer = vbDefault
             End If
          End If
     Case "Followbut"
         If tblbuttons(18) = 1 Then Exit Sub
         If tblbuttons(19) = 0 Then
            tblbuttons(19) = 1
            Toolbar1.Buttons(19).value = tbrPressed
            skyleftjump = True
         Else
            tblbuttons(19) = 0
            Toolbar1.Buttons(19).value = tbrUnpressed
            skyleftjump = False
            Exit Sub
            End If

'        If skyleftterracheck.Value = vbUnchecked Then
'           skyleftjump = False
'           terwt = False
'           Picture4.SetFocus
'        Else
        '      response = MsgBox("Do you want to hug the ground during your journey? " + _
                      "This will use a lot of system resources!", vbQuestion + vbYesNo, "SkyLight")
        '      If response = vbYes Then
        '         terwt = True
        '         End If

           lResult = FindWindow(vbNullString, terranam$)
        '   If lResult > 0 Then 'remove topmost status from terraviewer in order to display message
        '     ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
        '     End If
        '   response = MsgBox("If you want to hug the ground after the first jump, " + _
        '                     "keep the SHIFT-X key pressed continuously" + _
        '                     "during subsequent jumps until the buffer fills. " + _
        '                     "At this point, the altitude will be minimized " + _
        '                     "automatically.  Be warned that this will use a " + _
        '                     "lot of system resources!  " + _
        '                     "Happy journey!", vbInformation + vbOKOnly, "SkyLight")
        '   terwt = True
           'lResult = FindWindow(vbNullString, terranam$)
           If lResult > 0 Then 'restore topmost status
              ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
              'ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
           Else
              mapPictureform.mapPicture.SetFocus
              End If
           'End If
     Case "travelbut"
        'On Error GoTo errtravel
        If tblbuttons(20) = 0 Then
           tblbuttons(20) = 1
           travelmode = True
           travelnum% = 0
           Toolbar1.Buttons(20).value = tbrPressed
           showroute = True
        Else
           tblbuttons(20) = 0
           showroute = False
           travelmode = False
           Toolbar1.Buttons(20).value = tbrUnpressed
           If travelnum% >= 1 Then
               'process recorded values
tr50:           ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                If speed = 80 Then speed = 3000
                oldspeed = speed
                speed = Val(InputBox("Speed (mi/hr)", "Maps & More", speed, 1000, 2000))
                If speed = 0 Or speed = sEmpty Then speed = oldspeed
                On Error GoTo traverrorhand
                If appendtravel = True Then
                   savfile$ = appendfile$
                   GoTo tr75
                   End If
                CommonDialog1.CancelError = True
                If world = False Then
                   CommonDialog1.Filter = "Temporay travel files (*.trf)|*.trf|"
                   CommonDialog1.FilterIndex = 1
                   CommonDialog1.FileName = terradir$ + "\*.trf"
                Else
                   CommonDialog1.Filter = "world travel files (*.wtf)|*.wtf|"
                   CommonDialog1.FilterIndex = 1
                   CommonDialog1.FileName = "c:\dtm\*.wtf"
                   End If
                CommonDialog1.ShowSave
                'read the files coordinates, goto there, and then plot the obstructions
                savfile$ = CommonDialog1.FileName
tr75:           savfilnum% = FreeFile
                ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                myfile = Dir(savfile$)
                If myfile <> sEmpty And appendtravel = False Then
                   ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                   response = MsgBox("File already exists, do you want to overwrite it?", vbYesNo + vbExclamation + vbDefaultButton2, "Maps & More")
                   ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                   If response = vbNo Then GoTo tr50
                   End If
                Screen.MousePointer = vbHourglass
                Open savfile$ For Output As #savfilnum%
                If world = False Then
                    Print #savfilnum%, "[SETTINGS]"
                    Print #savfilnum%, "Use Route Elevation = 0"
                    Print #savfilnum%, "Use Route Speed = 0"
                    Print #savfilnum%, "Use Route Camera Offset = 0"
                    Print #savfilnum%, "Change Flight Pitch = 0"
                    Print #savfilnum%, "Round Trip = 0"
                    Print #savfilnum%, "Follow Ground = 1"
                    Print #savfilnum%, ""
                    Print #savfilnum%, "[POINTS], Speed = " + LTrim$(RTrim$(Str$(speed)))
                    Print #savfilnum%, ";  X  Y    Height   Ground Height   Pitch   Yaw   Roll   Speed   Turn Speed   Speed Accel   Camera Yaw   Camera Pitch   Time from prev point   Label"

                    Print #savfilnum%, "WayPoint = " + LTrim$(RTrim$(Str(travel(1, 1)))) + " " + LTrim$(RTrim$(travel(2, 1))) + " 0.000 0.000 0.000 0.000 0.000 0.000 0.000 0.000 0.000 0.000 0.000 10.000 #NOLABEL#"
                    For i% = 2 To travelnum%
                        xdif = travel(1, i%) - travel(1, i% - 1)
                        ydif = travel(2, i%) - travel(2, i% - 1)
                        calcspeed = Sqr(xdif ^ 2 + ydif ^ 2) * (0.5) / speed
                        If ydif >= 0 Then
                           If xdif >= 0 Then
                              yaw = 270 + (Atn(ydif / xdif)) / cd
                           ElseIf xdif < 0 Then
                              yaw = 90 + (Atn(ydif / xdif)) / cd
                              End If
                        ElseIf ydif < 0 Then
                           If xdif >= 0 Then
                              yaw = 270 + (Atn(ydif / xdif)) / cd
                           ElseIf xdif < 0 Then
                              yaw = 90 + (Atn(ydif / xdif)) / cd
                              End If
                           End If

                        yawstr = LTrim$(RTrim$(Format(Str$(yaw), "##0.000")))
                        Print #savfilnum%, "WayPoint = " + LTrim$(RTrim$(Str(travel(1, i%)))) + " " + LTrim$(RTrim$(travel(2, i%))) + " 0.000 0.000 0.000 " + yawstr + " 0.000 0.000 0.000 0.000 0.000 0.000 0.000 " + LTrim$(RTrim$(Format(Str$(calcspeed), "#####0.000"))) + " #NOLABEL#"

                    Next i%
                ElseIf world = True Then
                    Print #savfilnum%, "Number of route points"
                    Print #savfilnum%, travelnum%
                    Print #savfilnum%, "Speed"
                    Print #savfilnum%, speed
                    For i% = 1 To travelnum%
                       Print #savfilnum%, i%
                       Print #savfilnum%, travel(1, i%)
                       Print #savfilnum%, travel(2, i%)
                    Next i%
                    End If
                Close #savfilnum%
                Screen.MousePointer = vbDefault
                appendtravel = False
              End If
           End If
traverrorhand:
           Close
           Screen.MousePointer = vbHourglass
           Call blitpictures
           Screen.MousePointer = vbDefault
           ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
           Exit Sub
     Case "pausebut"
        If tblbuttons(23) = 0 Then
            Maps.Timer2.Enabled = False
            If world = True And showroute = True Then
               'send message to 3D Viwer to pause
               lResult = FindWindow(vbNullString, "3D Viewer")
               If lResult <> 0 Then
                  nmsg = SendMessage(lResult, WM_COMMAND, 1003, 0)
                  End If
               tblbuttons(23) = 1
               Toolbar1.Buttons(23).value = tbrPressed
               'Toolbar1.Buttons(18).Value = tbrUnpressed
               'Toolbar1.Buttons(17).Enabled = True
               'tblbuttons(18) = 0
               Exit Sub
               End If
            Screen.MouseIcon = LoadPicture("")
            Screen.MousePointer = vbDefault
            lResult = FindWindow(vbNullString, terranam$)
            ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
            Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
            Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
            'Call keybd_event(Asc("T"), 0, 0, 0)  'goes into tools menu
            'Call keybd_event(Asc("T"), 0, KEYEVENTF_KEYUP, 0)
            Call keybd_event(Asc("R"), 0, 0, 0)    'calls for the Route editor
            Call keybd_event(Asc("R"), 0, KEYEVENTF_KEYUP, 0)
            Call keybd_event(Asc("A"), 0, 0, 0)    'pauses the Route editor
            Call keybd_event(Asc("A"), 0, KEYEVENTF_KEYUP, 0)
            tblbuttons(23) = 1
            Toolbar1.Buttons(23).value = tbrPressed
            Toolbar1.Buttons(18).value = tbrUnpressed
            Toolbar1.Buttons(17).Enabled = True
            tblbuttons(18) = 0
         Else 'restart the travel file
           If world = True And showroute = True Then
              'send message to 3D Viwer to restart
              lResult = FindWindow(vbNullString, "3D Viewer")
              If lResult <> 0 Then
                 nmsg = SendMessage(lResult, WM_COMMAND, 1004, 0)
                 End If
               tblbuttons(23) = 0
               'Toolbar1.Buttons(18).Value = tbrPressed
               'tblbuttons(18) = 1
               Toolbar1.Buttons(17).Enabled = False
               Toolbar1.Buttons(23).value = tbrUnpressed
               Maps.Timer2.Enabled = True
               Exit Sub
               End If
            lResult = FindWindow(vbNullString, terranam$)
            ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
            Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
            Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
            'Call keybd_event(Asc("T"), 0, 0, 0)  'goes into tools menu
            'Call keybd_event(Asc("T"), 0, KEYEVENTF_KEYUP, 0)
            Call keybd_event(Asc("R"), 0, 0, 0)    'calls for the Route editor
            Call keybd_event(Asc("R"), 0, KEYEVENTF_KEYUP, 0)
            Call keybd_event(Asc("A"), 0, 0, 0)    'restarts the Route editor
            Call keybd_event(Asc("A"), 0, KEYEVENTF_KEYUP, 0)
            tblbuttons(23) = 0
            Toolbar1.Buttons(18).value = tbrPressed
            tblbuttons(18) = 1
            Toolbar1.Buttons(17).Enabled = False
            'Maps.Timer2.Interval = 3000
            Toolbar1.Buttons(23).value = tbrUnpressed
            Maps.Timer2.Enabled = True
            End If
     Case "stopbut"
        Maps.Timer2.Enabled = False
        If world = True Then GoTo to550
        Screen.MouseIcon = LoadPicture("")
        Screen.MousePointer = vbDefault
        lResult = FindWindow(vbNullString, terranam$)
        ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
        Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
        Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
        'Call keybd_event(Asc("T"), 0, 0, 0)  'goes into tools menu
        'Call keybd_event(Asc("T"), 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(Asc("R"), 0, 0, 0)    'calls for the Route editor
        Call keybd_event(Asc("R"), 0, KEYEVENTF_KEYUP, 0)
        Call keybd_event(Asc("T"), 0, 0, 0)    'stops the Route editor
        Call keybd_event(Asc("T"), 0, KEYEVENTF_KEYUP, 0)
to550:  If world = True And showroute = True Then
           'send message to 3D Viwer to stop
           lResult = FindWindow(vbNullString, "3D Viewer")
           If lResult <> 0 Then
              nmsg = SendMessage(lResult, WM_COMMAND, 1005, 0)
              End If
           End If
        Toolbar1.Buttons(18).value = tbrUnpressed
        Toolbar1.Buttons(23).value = tbrUnpressed
        Toolbar1.Buttons(23).Enabled = False
        Toolbar1.Buttons(24).Enabled = False
        Toolbar1.Buttons(17).Enabled = True
        tblbuttons(18) = 0
        tblbuttons(23) = 0
        tblbuttons(24) = 0
        Toolbar1.Buttons(25).value = tbrUnpressed
        Toolbar1.Buttons(25).Enabled = False
        tblbuttons(25) = 0
        showroute = False
        routnum% = 0
        routenum% = 0
        routeload = False
        Call blitpictures
     Case "showroutebut"
        If tblbuttons(25) = 0 Then
           tblbuttons(25) = 1
           Toolbar1.Buttons(25).value = tbrPressed
           showroute = True
           'openfilnum% = FreeFile 'open route file, and keep it open
           'Open openfile$ For Input As #openfilnum%
           Screen.MousePointer = vbHourglass
           Call blitpictures
           Screen.MousePointer = vbDefault
        Else
           'Close #openfilnum%
           Toolbar1.Buttons(25).value = tbrUnpressed
           tblbuttons(25) = 0
           showroute = False
           Screen.MousePointer = vbHourglass
           Call blitpictures
           Screen.MousePointer = vbDefault
           End If
     Case "sunrisekey"
        If graphwind = True Then
           If mapgraphfm.Caption = "Sunrise horizon profile" Then
              ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
              'Exit Sub
              End If
           End If

        If noheights = True Then
           On Error GoTo sunerr
           If world = True Then
                'check if there is a USGUS EROS CD or SRTM CD in the CD-drive
                 myfile = Dir(worlddtm + ":\Gt30dem.gif")
                If myfile = sEmpty Then
                   'check if there are stored DTM files in c:\dtm
                   doclin$ = Dir("c:\dtm\*.BIN")
                   myfile = Dir("c:\dtm\eros.tm3")
                   If doclin$ <> sEmpty And myfile <> sEmpty And Dir("c:\dtm\*.BI1") <> sEmpty Then
                     'leave rest of checking for sunrisesunset routine
                     checkdtm = True
                     Call sunrisesunset(1)
                   ElseIf Dir(srtmdtm & ":\3AS\", vbDirectory) <> sEmpty Or _
                          Dir(srtmdtm & ":\USA\", vbDirectory) <> sEmpty Then
                     checkdtm = True
                     Call sunrisesunset(1)
                   ElseIf Not NoCDWarning Then
                      Maps.Toolbar1.Buttons(26).value = tbrUnpressed
                      ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                      response = MsgBox("USGS EROS CD not found!  Please enter the appropriate CD, and then press the DTM button!", vbCritical + vbOKOnly, "Maps & More")
                      ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                      NoCDWarning = True
                      Exit Sub
                   End If
                Else
                   noheights = False
                   Maps.Toolbar1.Buttons(1).value = tbrPressed
                   tblbuttons(1) = 1
                End If
           Else
               On Error Resume Next
               myfile = sEmpty
               myfile = Dir(israeldtm + ":\dtm\dtm-map.loc")
               If myfile = sEmpty Then
                   tblbuttons(26) = 0
                   Toolbar1.Buttons(26).value = tbrUnpressed
                   MsgBox "DTM CD or data not found.  Place the DTM CD in the drive and try again", vbCritical + vbOKOnly, "Maps & More"
                   Exit Sub
                Else
                  noheights = False
                  Maps.Toolbar1.Buttons(1).value = tbrPressed
                  tblbuttons(1) = 1
                End If
           End If
        End If
        If Maps.Toolbar1.Buttons(26).value = tbrPressed Then
           If world = True Then
              Call sunrisesunset(1)
           Else
              Call EYsunrisesunset(1)
           End If
        End If
     Case "sunsetkey"
        If graphwind = True Then
           If mapgraphfm.Caption = "Sunset horizon profile" Then
              ret = SetWindowPos(mapgraphfm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
              'Exit Sub
              End If
           End If

        If noheights = True Then
            If world = True Then
                 'check if there is a USGUS EROS CD in the CD-drive
                 On Error GoTo sunerr
                 myfile = Dir(worlddtm + ":\Gt30dem.gif")
                 If myfile = sEmpty Then
                    'check if there are stored DTM files in c:\dtm
                    doclin$ = Dir("c:\dtm\*.BIN")
                    myfile = Dir("c:\dtm\eros.tm3")
                    If doclin$ <> sEmpty And myfile <> sEmpty And Dir("c:\dtm\*.BI1") <> sEmpty Then
                      'leave rest of checking for sunrisesunset routine
                      checkdtm = True
                      Call sunrisesunset(0)
                    ElseIf Not NoCDWarning Then
                      Maps.Toolbar1.Buttons(26).value = tbrUnpressed
                      ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                      response = MsgBox("USGS EROS CD not found!  Please enter the appropriate CD, and then press the DTM button!", vbCritical + vbOKOnly, "Maps & More")
                      ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
                      NoCDWarning = True
                      Exit Sub
                    End If
                 Else
                    noheights = False
                    Maps.Toolbar1.Buttons(1).value = tbrPressed
                    tblbuttons(1) = 1
                    End If
            Else
               On Error Resume Next
               myfile = sEmpty
               myfile = Dir(israeldtm + ":\dtm\dtm-map.loc")
               If myfile = sEmpty Then
                   tblbuttons(27) = 0
                   Toolbar1.Buttons(27).value = tbrUnpressed
                   MsgBox "DTM CD or data not found.  Place the DTM CD in the drive and try again", vbCritical + vbOKOnly, "Maps & More"
                   Exit Sub
               Else
                  noheights = False
                  Maps.Toolbar1.Buttons(1).value = tbrPressed
                  tblbuttons(1) = 1
                End If
             End If
          End If
        If Maps.Toolbar1.Buttons(27).value = tbrPressed Then
           If world = True Then
              Call sunrisesunset(0)
           Else
              Call EYsunrisesunset(0)
           End If
        End If
        
     Case "Tempbut"
        If mapTempfrm.Visible = False Then
           mapTempfrm.Visible = True
           ret = SetWindowPos(mapTempfrm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
        Else
           ret = SetWindowPos(mapTempfrm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
           End If
     Case Else
   End Select
   
   For i% = 1 To Toolbar1.Buttons.count
       If tblbuttons(i%) = 1 Then
          Toolbar1.Buttons(i%).value = tbrPressed
       Else
          Toolbar1.Buttons(i%).value = tbrUnpressed
       End If
   Next i%
   Toolbar1.Refresh
   
Exit Sub
errtravel:
   Exit Sub
sunerr:
   myfile = sEmpty
   Resume Next

   On Error GoTo 0
   Exit Sub

Toolbar1_ButtonClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Toolbar1_ButtonClick of Form Maps"
End Sub
Private Sub MDIform_queryunload(Cancel As Integer, UnloadMode As Integer)
    If Forms.count > 2 Then
       For i% = 0 To Forms.count - 1
          ret = SetWindowPos(Forms(i%).hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
       Next i%
       response = MsgBox("Exit Maps & More?", vbQuestion + vbYesNoCancel + vbMsgBoxSetForeground, "Maps & More Exit")
       If response <> vbYes Then
          Cancel = True
          For i% = 0 To Forms.count - 1
             ret = SetWindowPos(Forms(i%).hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
          Next i%
          Exit Sub
          End If
       End If

    'If world = False Then
    '   lResult = FindWindow(vbNullString, terranam$)
    '   If lResult > 0 Then
    '      ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
    '      End If
    'ElseIf world = True Then
    '   lResult = FindWindow(vbNullString, "3D Viewer")
    '   If lResult > 0 Then
    '      ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
    '      End If
    '   End If
    'If mapPictureform.Visible = True Then
    '   ret = SetWindowPos(mapPictureform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
    '   End If
    'response = MsgBox("Are you sure you wan't to quit?", vbQuestion + vbYesNo, "Maps & More")
    'If response = vbNo Then
    '   Exit Sub
    '   End If

    Maps.Timer3.Enabled = False
    'response = MsgBox("reboot?", vbQuestion + vbOKCancel) '<<<<<<<
    'If response = vbOK Then reboot = True
    If exit1 = True And exit2 = False Then
       exit2 = True
    ElseIf exit1 = True And exit2 = True Then
       GoTo fq100
       End If
    If reboot = True Then exit2 = True
    Close 'first close any open files

    'now record last kmxc,kmyc,hgt in save file
    filnum% = FreeFile
    Open drivjk$ + "mapposition.sav" For Output As #filnum%
    'If hgtpos = sEmpty Then hgtpos = 0
    Write #filnum%, kmxsky, kmysky, hgtpos
    If Maps.Text7.Text = sEmpty Then
       hgtworld = 0
    Else
       hgtworld = Maps.Text7.Text
       End If
    Write #filnum%, lon, lat, hgtworld
    Write #filnum%, maxangf%, diflogf%, diflatf%, fullrangef%, viewmodef%, modevalf
    Write #filnum%, DTMflag
    Write #filnum%, maxangfs%, diflogfs%, diflatfs%, fullrangefs%, viewmodefs%, modevalfs
    Write #filnum%, CalculateProfile
    Write #filnum%, AziStepf%
    Write #filnum%, rderos2_use
    Write #filnum%, IgnoreTiles%
    Write #filnum%, autoazirange%
    Write #filnum%, TemperatureModel%
    Close #filnum%

   'shut off terraviewer if still activated
    If terranam$ = sEmpty Then GoTo fq4
    lResult = FindWindow(vbNullString, terranam$)
    If lResult > 0 Then
       ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
       Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
       Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
       Call keybd_event(Asc("F"), 0, 0, 0)  'goes into Settings menu
       Call keybd_event(Asc("F"), 0, KEYEVENTF_KEYUP, 0)
       Call keybd_event(Asc("X"), 0, 0, 0)  'goes into Settings menu
       Call keybd_event(Asc("X"), 0, KEYEVENTF_KEYUP, 0)
       Maps.Toolbar1.Buttons(17).value = tbrUnpressed
       End If

   'shut off 3D Viewer if still activated
    lResult = FindWindow(vbNullString, "3D Viewer")
    If lResult > 0 Then
       ret = BringWindowToTop(lResult) 'bring TerraViewer to top of Z order
       Call keybd_event(VK_F10, 0, 0, 0)  'activates alt key
       Call keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, 0)
       Call keybd_event(Asc("F"), 0, 0, 0)  'goes into Files menu
       Call keybd_event(Asc("F"), 0, KEYEVENTF_KEYUP, 0)
       Call keybd_event(Asc("X"), 0, 0, 0)  'Exit
       Call keybd_event(Asc("X"), 0, KEYEVENTF_KEYUP, 0)
       Maps.Toolbar1.Buttons(26).value = tbrUnpressed
       Maps.Toolbar1.Buttons(27).value = tbrUnpressed
       End If


'now save the world DTM extraction if desired
fq4: doclin$ = Dir(ramdrive + ":\*.bin")
   myfile = Dir(drivjk$ + "eros.tm3")
   myfile2 = Dir("c:\dtm\eros.tm3")
   If myfile2 = sEmpty Then GoTo fq02
   If doclin$ <> sEmpty And myfile <> sEmpty And Dir("c:\dtm\*.bi1") <> sEmpty Then
      'check if .BIN file has already been saved
       If myfile2 <> sEmpty Then
          filtmp% = FreeFile
          Open "c:\dtm\eros.tm3" For Input As #filtmp%
          Line Input #filtmp%, filn33$
          Input #filtmp%, L1ch, L2ch, hgtch, angch, apch, modch%, modvalch%
          Input #filtmp%, beglogch, endlogch, beglatch, endlatch
          Close #filtmp%
          filtmp% = FreeFile
          Open drivjk$ + "eros.tm3" For Input As #filtmp%
          Line Input #filtmp%, filn33$
          Input #filtmp%, l1, l2, hgt, angch, apch, Mode%, modval%
          Input #filtmp%, beglog, endlog, beglat, endlat
          Close #filtmp%
          If modch% = Mode% And modval% = modvalch% And _
             ((Mode% = 1 And Abs(endlogch - endlog) < 0.05 And beglog - beglogch >= -0.0001) Or _
             (Mode% = 0 And endlogch - endlog >= -0.0001 And Abs(beglogch - beglog) < 0.05)) And _
             Abs(beglatch - beglat) < 0.05 And Abs(endlatch - endlat) < 0.05 Then
             'its the same file, so don't save it again
             GoTo fq5
          Else 'old file found saved on c:\dtm, delete it since it is not relevant
             myfile3 = Dir("c:\dtm\*.bin")
             If myfile3 <> sEmpty Then Kill "c:\dtm\" + myfile3
             myfile3 = Dir("c:\dtm\*.bi1")
             If myfile3 <> sEmpty Then Kill "c:\dtm\" + myfile3
             Kill "c:\dtm\" + myfile2
             End If
          End If

fq02: lResult = FindWindow(vbNullString, terranam$)
      If lResult <> 0 Then ret = BringWindowToTop(lResult)
      ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
      
      myfile = Dir("c:\dtm\*.tBI")
      If myfile <> sEmpty Then

      response = MsgBox("Before exiting do you want to save the DTM? (*.BIN/*.BI1) files?", vbQuestion + vbYesNoCancel, "Maps & More")
      If response = vbYes Then
'         myfile = Dir("c:\dtm\" + doclin$)
'         If myfile <> sEmpty Then
'            response = MsgBox("File already exists, overwrite it?", vbExclamation + vbYesNo, "Maps & More")
'            If response = vbNo Then
'               GoTo fq5
'               End If
'            End If
         myfile = Dir("c:\dtm\*.tBI")
         If myfile <> sEmpty Then
            FileCopy "c:\dtm\" + myfile, "c:\dtm\" + Mid$(myfile, 1, Len(myfile) - 4) + ".BI1"
            Kill "c:\dtm\" + myfile
         Else
            response = MsgBox("Sorry can't save the DTM since the BI1 file was not found!", vbExclamation + vbOKOnly, "Maps & More")
            GoTo fqq5
            End If
         FileCopy ramdrive + ":\" + doclin$, "c:\dtm\" + doclin$
         FileCopy drivjk$ + "eros.tm3", "c:\dtm\eros.tm3"
         If Dir(ramdrive + ":\land.x") <> sEmpty And Dir(ramdrive + ":\land.tm3") <> sEmpty Then
            response = MsgBox("Do you wan't to save the 3D Viewer (*.x) file and (*.tm3) file?", vbQuestion + vbYesNoCancel, "Maps & More")
            If response = vbYes Then
               FileCopy ramdrive + ":\land.x", "c:\dtm\land.x"
               FileCopy ramdrive + ":\land.tm3", "c:\dtm\land.tm3"
            ElseIf response = vbCancel Then
               Cancel = 1
               Exit Sub
               End If
            End If
     ElseIf response = vbNo Then 'erase temporarily saved .BI1 file
         myfile = Dir("c:\dtm\*.tBI")
         If myfile <> sEmpty Then
            Kill "c:\dtm\" + myfile
            End If
     ElseIf response = vbCancel Then
         Cancel = 1
         Exit Sub
         End If
         
     End If

fqq5: dy1 = -300
      dx1 = 300
      If exit1 = True Then dx1 = -300
      Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
      End If

   'then unload any hidden forms
fq5: Maps.Visible = False
   On Error GoTo fq10
    For i% = 1 To Forms.count - 1
       Unload Forms(i%)
    Next i%
    'now unload main form
fq10: Unload Maps
    Set Maps = Nothing
'now restore task bar
  'GoTo 999
  'set Window's taskbar to AutoHide
   'waitime = Timer
   'Do Until Timer > waitime + 0.1
   '   DoEvents
   'Loop
fq100:
   GoTo map900
   
   'skip the old mouse way of making taskbar reappear

'   dx1 = 0
'   If exit2 = True Or exit3 = True Then 'And reboot = False Then
'      dx1 = 245
'      End If
'   dy1 = 600 '300
'   If reboot = True Then
'      dy1 = 300
'      dx1 = -300
'      Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'      dy1 = 0
'      dx1 = 250
'      End If
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   'If reboot = True Then
'   '   waitime = Timer
'   '   Do Until Timer > waitime + 0.1
'   '      DoEvents
'   '   Loop
'   '   dx1 = 5
'   '   dy1 = -5
'   '   Call mouse_event(MOUSEEVENTF_MOVE, adx1*dx1, bdy1*dy1, 0, 0)  'move mouse to Location item
'   '   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'   '   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
'   '   waitime = Timer
'   '   Do Until Timer > waitime + 0.1
'   '      DoEvents
'   '   Loop
'   '   Call keybd_event(VK_UP, 0, 0, 0)
'   '   Call keybd_event(VK_UP, 0, KEYEVENTF_KEYUP, 0)
'   '   Call keybd_event(VK_RETURN, 0, 0, 0) 'enters return
'   '   Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
'   '   waitime = Timer
'   '   Do Until Timer > waitime + 0.1
'   '      DoEvents
'   '   Loop
'   '   Call keybd_event(VK_DOWN, 0, 0, 0)
'   '   Call keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0)
'   '   Call keybd_event(VK_RETURN, 0, 0, 0) 'enters return
'   '   Call keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
'   '   End
'   '   End If
'   waitime = Timer
'   Do Until Timer > waitime + 0.5 ' 0.1 <---changed
'      DoEvents
'   Loop
'   Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'   Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0) 'move mouse to Location item
'   waitime = Timer
'   Do Until Timer > waitime + 0.5 '0.1 <---changed
'      DoEvents
'   Loop
'   dx1 = -20
'   If exit2 = True Or exit3 = True Then
'      dx1 = 20
'      End If
'   dy1 = -5
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
'   dx1 = 0
'   dy1 = -65
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   waitime = Timer
'   Do Until Timer > waitime + 0.5 ' 0.1 <---changed
'      DoEvents
'   Loop
'   dx1 = -1500
'   dy1 = 0
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   'waitime = Timer
'   'Do Until Timer > waitime + 3 '<<<<<<
'   '   DoEvents
'   'Loop
'   dx1 = 30
'   dy1 = 0
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   waitime = Timer
'   Do Until Timer > waitime + 0.5 '0.1 <---changed
'      DoEvents
'   Loop
'   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
'   'waitime = Timer
'   'Do Until Timer > waitime + 3 '<<<<<
'   '   DoEvents
'   'Loop
'   dx1 = 50
'   dy1 = 55
'   Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item
'   waitime = Timer
'   Do Until Timer > waitime + 0.5 '0 1 <---changed
'      DoEvents
'   Loop
'   Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'move mouse to Location item
'   Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'move mouse to Location item
'   waitime = Timer
'   Do Until Timer > waitime + 1
'      DoEvents
'   Loop
   
map900:
   If reboot = True Then
      'copy maps&more lnk file to the START directory
      SourceFile = drivjk$ + "maps&m~1.lnk"
      DestinationFile = "c:\windows\startm~1\programs\startup\maps&m~1.lnk"
      FileCopy SourceFile, DestinationFile
      ret = ExitWindowsEx(EWX_REBOOT, dwReserved)
      End If

999   End
9999
End Sub


Private Sub userspeedfm_Click()
Dim lResult As Long
userspeedfm.Checked = False
lResult = FindWindow(vbNullString, terranam$)
If lResult > 0 Then ret = SetWindowPos(lResult, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
If mapPictureform.Visible = True Then ret = SetWindowPos(mapPictureform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
oldspeed = speed
10 speed = Val(InputBox("Speed (mi/hr)", "Maps & More", speed, 1000, 2000))
   If speed < 0 Then
      response = MsgBox("Not a valid speed!", vbOKOnly + vbCritical, "Maps & More")
      GoTo 10
      End If
   If speed = 0 Then
      speed = oldspeed
      GoTo 50
      End If
   If speed > 120 And world = False Then
      response = MsgBox("Warning, this speed may be too fast for the recorded point density (in which case the TerraViewer will deviate from the recorded route).  Do you still want to use this high speed?", vbYesNo + vbQuestion, "Maps & More")
      If response = vbNo Then
         GoTo 10
         End If
      End If
50 userspeedfm.Checked = True
   mihr50fm.Checked = False
   mihr60fm.Checked = False
   mihr70fm.Checked = False
   mihr80fm.Checked = False
   mihr90fm.Checked = False
   mihr100fm.Checked = False
   mihr110fm.Checked = False
   mihr120fm.Checked = False
   speeddefaultfm.Checked = False

 dx1 = 0
 dy1 = -150 'move the pointer away from terraviewer window
 Call mouse_event(MOUSEEVENTF_MOVE, adx1 * dx1, bdy1 * dy1, 0, 0) 'move mouse to Location item

If lResult > 0 Then ret = SetWindowPos(lResult, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
If mapPictureform.Visible = True Then ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub
Private Sub picture4_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub
Private Sub toolbar1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub
Private Sub picture1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub
Private Sub statusbar1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lHelp As Long
   Select Case KeyCode
      Case vbKeyF1
         Maps.CommonDialog2.HelpFile = "Maps&More.hlp"
         Maps.CommonDialog2.HelpCommand = cdlHelpContents
         Maps.CommonDialog2.ShowHelp
         waitime = Timer
         Do Until Timer > waitime + 5
            DoEvents
         Loop
         lHelp = FindWindow(vbNullString, "Maps & More Help")
         If lHelp > 0 Then
            ret = SetWindowPos(lHelp, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
            End If
      Case Else
   End Select
End Sub

Private Sub map50butsub()
       If mapPictureform.Visible = False Or map400 = True Or world = True Then
          topofm.Enabled = True
          openbatfm.Enabled = True
          If map400 = True Then
             map400 = False
             Toolbar1.Buttons(6).value = tbrUnpressed
             tblbuttons(6) = 0
             End If
          If world = True Then
             world = False
             mapPictureform.Visible = False
             mapPictureform.Width = sizex + 60 '60 is the size (pixels) of the borders
             mapPictureform.Height = sizey + 60
             mapPictureform.mapPicture.Width = sizex
             mapPictureform.mapPicture.Height = sizey
             mapxdif = mapPictureform.Width - mapPictureform.mapPicture.Width
             mapydif = mapPictureform.Height - mapPictureform.mapPicture.Height
             If world = True Then
               mapxdif = mapxdif + 35
               mapydif = mapydif + 35
               End If
             mapwi = mapPictureform.Width
             maphi = mapPictureform.Height
             tblbuttons(3) = 0
             tblbuttons(8) = 0
             Maps.Toolbar1.Buttons(3).value = tbrUnpressed
             Maps.Toolbar1.Buttons(3).Enabled = False
             Toolbar1.Buttons(8).value = tbrUnpressed
             'Check if there is enough RAM memory to run
             'Windows version of rdhal.bat
             If mapEROSDTMwarn.Visible = True Then
               Unload mapEROSDTMwarn
               End If
             End If
          For i% = 2 To 15
             Toolbar1.Buttons(i%).Enabled = True
          Next i%
          Toolbar1.Buttons(26).value = tbrUnpressed
          Toolbar1.Buttons(27).value = tbrUnpressed
          If RdHalYes Then 'enable sunrise/sunset calculations
            Toolbar1.Buttons(26).Enabled = True
            Toolbar1.Buttons(27).Enabled = True
          Else
            Toolbar1.Buttons(26).Enabled = False
            Toolbar1.Buttons(27).Enabled = False
          End If
          If noheights = False Then
             mnuCrossSection.Enabled = True
             mnuFirstPoint.Enabled = True
             mnuSecondPoint.Enabled = True
             End If
          appendfrm.Enabled = True
          Toolbar1.Buttons(20).Enabled = True
          Toolbar1.Buttons(21).Enabled = True
          searchfm.Enabled = True
          Combo1.Enabled = True
          coordmode% = 1
          Maps.Text4.Visible = False
          Maps.Label4.Visible = False
          coordmode2% = 1
          Label1.Caption = "ITMx"
          Label2.Caption = "ITMy"
          Label5.Caption = "ITMx"
          Label6.Caption = "ITMy"
          Maps.Text1.Text = "0"
          Maps.Text2.Text = "0"
          Maps.Text3.Text = "0"
          Maps.Text5.Text = kmxc
          Maps.Text6.Text = kmyc
          If topotype% = 1 Then
             Maps.Text5 = Fix(kmxc - 5785 * km50x + 0.5) '***********
             Maps.Text6 = Fix(kmyc + 10615 * km50y + 0.5)
             End If
          Maps.Text7.Text = hgt50c
          If Maps.Text7.Text = sEmpty Then
             hgt50c = 0: hgtpos = 0
             Maps.Text7.Text = "0"
             End If
          Picture4.Visible = True
          lResult = FindWindow(vbNullString, terranam$)
          If lResult > 0 And terranam$ <> sEmpty Then
             For i% = 18 To 21
                Toolbar1.Buttons(i%).Enabled = True
             Next i%
             Loadfm.Enabled = True
             If Dir(ramdrive + ":\travlog.x") <> sEmpty Then recoverroutefm.Enabled = True
             End If
          map50 = True
          Toolbar1.Buttons(7).value = tbrPressed
          tblbuttons(7) = 1
          Toolbar1.Buttons(17).Enabled = True
          Call loadpictures  'load appropriate map tiles into off-screen buffers
          Call blitpictures   'blit desired portions of the off-screen buffers to the screen
          If kmx50c = 0 And kmy50c = 0 And kmx400c = 0 And kmy400c = 0 Then
             kmx50c = kmxc: kmy50c = kmyc: hgt50c = hgt
             kmx400c = kmxc: kmy400c = kmyc: hgt400c = hgt
             ''convert to sky coordinates
             'mode% = 1
             'kmxo = kmxc: kmyo = kmyc
             kmxsky = kmxc: kmysky = kmyc
             'Call ITMSKY(kmxo, kmyo, T1, T2, mode%)
             'Maps.Text5.Text = T1
             'Maps.Text6.Text = T2
             'Maps.Text7.Text = hgtpos
             End If
          If mapPictureform.Visible = True Then
             ret = SetWindowPos(mapPictureform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
             End If
       Else
          map50 = False
          openbatfm.Enabled = False
          topofm.Enabled = False
          Toolbar1.Buttons(7).value = tbrUnpressed
          tblbuttons(7) = 0
          mapPictureform.Visible = False
          appendfrm.Enabled = False
          For i% = 9 To 15
             Toolbar1.Buttons(i%).Enabled = False
             tblbuttons(i%) = 0
          Next i%
          For i% = 18 To 27 '24 <<<changed, needs checking
             Toolbar1.Buttons(i%).Enabled = False
             tblbuttons(i%) = 0
          Next i%
          Toolbar1.Buttons(26).value = tbrUnpressed
          tblbuttons(26) = 0
          Toolbar1.Buttons(27).value = tbrUnpressed
          tblbuttons(26) = 0
          Toolbar1.Buttons(2).Enabled = False
          Toolbar1.Buttons(4).Enabled = False
          Toolbar1.Buttons(20).value = tbrUnpressed
          tblbuttons(20) = 0
          Loadfm.Enabled = False
          recoverroutefm.Enabled = False
          Combo1.Enabled = False
          Picture4.Visible = False
          searchfm.Enabled = False
          If tblbuttons(4) = 1 Then
             Maps.Toolbar1.Buttons(4).value = tbrUnpressed
             tblbuttons(4) = 0
             obstflag = False
             End If
          End If
End Sub

Private Sub MDIform_resize()
   'MDIform must always be maximized
   On Error GoTo 999
'   If init = True Then
'      init = False
'      cx = GetSystemMetrics(SM_CXSCREEN)
'      cy = GetSystemMetrics(SM_CYSCREEN)
'      init = True
'      ret = SetWindowPos(Maps.hwnd, HWND_TOP, 0, 0, cx, cy, SWP_SHOWWINDOW)
'   Else
      Maps.WindowState = vbNormal 'don't let the window state change
      Me.Left = -60
      Me.Top = -60
      Me.Width = Screen.Width + 120
      Me.Height = Screen.Height + 120
'      End If
999  Exit Sub
End Sub
Function GPS_connect()

   Load GPStest
   
End Function
Public Sub GPSInitialization()

   'initializes GPS communication

    GPS_timer_trials = 0 'first attempt to connect with the following connection values:
       
    GPSConnectString0 = GetSetting(App.Title, "Settings", "GPS serial-USB connection string")
    GPSConnectString = GPSConnectString0
    If GPSConnectString = sEmpty Then
       GPSConnectString = "38400,N,8,1" 'default baud rate, parity, data bit, stop bit
       End If
    GPS_connect
                
  
End Sub
