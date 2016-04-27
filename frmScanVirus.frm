VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmScanVirus 
   BackColor       =   &H80000012&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Beetle Av"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   14565
   Icon            =   "frmScanVirus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SuperAv.ShapeButton ShapeButton3 
      Height          =   3615
      Left            =   11280
      TabIndex        =   136
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6376
      ButtonStyle     =   8
      ButtonTheme     =   3
      BackColor       =   11903133
      BackColorPressed=   15525862
      BackColorHover  =   16250356
      BorderColor     =   11903133
      BorderColorPressed=   15525862
      BorderColorHover=   16250356
      Caption         =   ""
      Picture         =   "frmScanVirus.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SuperAv.ShapeButton ShapeButton2 
      Height          =   3615
      Left            =   8400
      TabIndex        =   135
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6376
      ButtonStyle     =   8
      ButtonStyleColors=   0
      ButtonTheme     =   3
      BackColor       =   11903133
      BackColorPressed=   15525862
      BackColorHover  =   16250356
      BorderColor     =   11903133
      BorderColorPressed=   15525862
      BorderColorHover=   16250356
      Caption         =   ""
      Picture         =   "frmScanVirus.frx":2CC0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SuperAv.ShapeButton ShapeButton1 
      Height          =   3615
      Left            =   5640
      TabIndex        =   134
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6376
      ButtonStyle     =   8
      ButtonStyleColors=   0
      ButtonTheme     =   3
      ButtonType      =   1
      BackColor       =   11903133
      BackColorPressed=   15525862
      BackColorHover  =   16250356
      BorderColor     =   11903133
      BorderColorPressed=   15525862
      BorderColorHover=   16250356
      Caption         =   ""
      Picture         =   "frmScanVirus.frx":474A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox sIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   123
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrFadeout 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8100
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   7560
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8550
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   8625
      Visible         =   0   'False
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2919
            MinWidth        =   2919
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2717
            MinWidth        =   2717
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3529
            MinWidth        =   3529
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "2:20 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      Height          =   6570
      Left            =   120
      Picture         =   "frmScanVirus.frx":61D1
      ScaleHeight     =   6570
      ScaleWidth      =   4095
      TabIndex        =   21
      Top             =   1920
      Width           =   4095
      Begin SuperAv.ShapeButton ShapeButton5 
         Height          =   375
         Left            =   1200
         TabIndex        =   138
         Top             =   6120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         ButtonStyle     =   0
         ButtonStyleColors=   0
         ButtonTheme     =   3
         BackColor       =   11903133
         BackColorPressed=   15525862
         BackColorHover  =   16250356
         BorderColor     =   11903133
         BorderColorPressed=   11903133
         BorderColorHover=   11903133
         Caption         =   ""
         Picture         =   "frmScanVirus.frx":7FDD
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
      Begin SuperAv.ShapeButton ShapeButton4 
         Height          =   1215
         Left            =   2280
         TabIndex        =   137
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         ButtonStyle     =   8
         ButtonStyleColors=   0
         ButtonTheme     =   3
         BackColor       =   11903133
         BackColorPressed=   15525862
         BackColorHover  =   16250356
         BorderColor     =   11903133
         BorderColorPressed=   15525862
         BorderColorHover=   16250356
         Caption         =   ""
         Picture         =   "frmScanVirus.frx":8503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdMenuProcess 
         Height          =   375
         Left            =   1080
         TabIndex        =   24
         Tag             =   "View Running Process"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ButtonStyle     =   7
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   " Processes"
         Picture         =   "frmScanVirus.frx":8E31
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdMenuStartup 
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Tag             =   "View Startup"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         ButtonStyle     =   7
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Startup"
         Picture         =   "frmScanVirus.frx":93CB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdMenuTweakReg 
         Height          =   375
         Left            =   720
         TabIndex        =   26
         Tag             =   "Repair Your Registry"
         Top             =   3840
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         ButtonStyle     =   7
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Tweak  Registry"
         Picture         =   "frmScanVirus.frx":9965
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdMenuOptions 
         Height          =   1215
         Left            =   2280
         TabIndex        =   27
         Tag             =   "Options"
         ToolTipText     =   "Start Scan"
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         ButtonStyle     =   8
         ButtonStyleColors=   0
         ButtonTheme     =   3
         PictureAlignment=   1
         BackColor       =   11903133
         BackColorPressed=   15525862
         BackColorHover  =   16250356
         BorderColor     =   11903133
         BorderColorPressed=   15525862
         BorderColorHover=   16250356
         Caption         =   ""
         Picture         =   "frmScanVirus.frx":9EFF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdQuarantine 
         Height          =   1215
         Left            =   600
         TabIndex        =   28
         Tag             =   "View Quarantine "
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         ButtonStyle     =   8
         ButtonStyleColors=   0
         ButtonTheme     =   3
         PictureAlignment=   1
         BackColor       =   11903133
         BackColorPressed=   15525862
         BackColorHover  =   16250356
         BorderColor     =   11903133
         BorderColorPressed=   15525862
         BorderColorHover=   16250356
         ForeColorPressed=   0
         ForeColorHover  =   0
         EffectColor     =   0
         Caption         =   ""
         Picture         =   "frmScanVirus.frx":AB82
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdMenuScan 
         Height          =   1215
         Left            =   600
         TabIndex        =   32
         Tag             =   "Home Scan"
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         ButtonStyle     =   8
         ButtonStyleColors=   0
         ButtonTheme     =   3
         PictureAlignment=   1
         BackColor       =   11903133
         BackColorPressed=   15525862
         BackColorHover  =   16250356
         BorderColor     =   11903133
         BorderColorPressed=   15525862
         BorderColorHover=   16250356
         Caption         =   ""
         Picture         =   "frmScanVirus.frx":B8E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdSignature 
         Height          =   735
         Left            =   -120
         TabIndex        =   121
         Tag             =   "Update Virus to DB"
         Top             =   5040
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         ButtonStyle     =   6
         ButtonStyleColors=   3
         PictureAlignment=   1
         BackColor       =   13819613
         BackColorPressed=   16777215
         BackColorHover  =   16777215
         ForeColor       =   16777215
         EffectColor     =   12582912
         Caption         =   "Update "
         Picture         =   "frmScanVirus.frx":C173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdExit 
         Height          =   375
         Left            =   2880
         TabIndex        =   122
         Tag             =   "Exit From Application"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         ButtonStyle     =   7
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Exit"
         Picture         =   "frmScanVirus.frx":C7F8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdSetAttrib 
         Height          =   390
         Left            =   2280
         TabIndex        =   124
         Top             =   3840
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   688
         ButtonStyle     =   7
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Set Attribute"
         Picture         =   "frmScanVirus.frx":CD92
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdSystray 
         Height          =   375
         Left            =   240
         TabIndex        =   125
         Tag             =   "Hide to System Tray"
         ToolTipText     =   "Start Scan"
         Top             =   6120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         ButtonStyle     =   0
         ButtonStyleColors=   0
         ButtonTheme     =   3
         PictureAlignment=   1
         BackColor       =   11903133
         BackColorPressed=   15525862
         BackColorHover  =   16250356
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   ""
         Picture         =   "frmScanVirus.frx":D32C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picTweak 
      BackColor       =   &H00FFFFFF&
      Height          =   6090
      Left            =   3720
      ScaleHeight     =   6030
      ScaleWidth      =   7980
      TabIndex        =   79
      Top             =   9240
      Width           =   8040
      Begin VB.Timer tmrFix 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5850
         Top             =   4425
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   150
         TabIndex        =   109
         Top             =   5025
         Width           =   7740
         Begin SuperAv.ProgressBar progFixReg 
            Height          =   390
            Left            =   150
            TabIndex        =   110
            Top             =   225
            Width           =   7440
            _ExtentX        =   13123
            _ExtentY        =   688
            Value           =   0
            Theme           =   8
            TextStyle       =   2
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "U11D ProgressBar"
            TextEffectColor =   16777215
            TextEffect      =   5
         End
      End
      Begin SuperAv.ShapeButton cmdApplyTweak 
         Height          =   390
         Left            =   150
         TabIndex        =   104
         Top             =   3900
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         ButtonStyle     =   7
         CaptionAlignment=   3
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   11907757
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Apply"
         Picture         =   "frmScanVirus.frx":D822
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame30 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display Properties Restrictions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1515
         Left            =   4485
         TabIndex        =   85
         Top             =   75
         Width           =   3390
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Display Settings Page "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   150
            TabIndex        =   89
            Tag             =   "NoDispSettingsPage"
            Top             =   1185
            Width           =   3105
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Screen Saver Settings Page "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   150
            TabIndex        =   88
            Tag             =   "NoDispScrSavPage"
            Top             =   915
            Width           =   3105
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Display Background Page "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   150
            TabIndex        =   87
            Tag             =   "NoDispBackgroundPage"
            Top             =   645
            Width           =   3105
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Display Appearance Page "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   86
            Tag             =   "NoDispAppearancePage"
            Top             =   375
            Width           =   3105
         End
      End
      Begin VB.Frame Frame31 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Windows Security Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1515
         Left            =   4485
         TabIndex        =   80
         Top             =   2175
         Width           =   3390
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable System Tray "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   20
            Left            =   150
            TabIndex        =   84
            Tag             =   "NoTrayItemsDisplay"
            Top             =   1095
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable context menus for the Taskbar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   19
            Left            =   150
            TabIndex        =   83
            Tag             =   "NoTrayContextMenu"
            Top             =   855
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Network Neighborhood Icon"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   18
            Left            =   150
            TabIndex        =   82
            Tag             =   "NoNetHood"
            Top             =   615
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable the Shut Down Command"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   17
            Left            =   150
            TabIndex        =   81
            Tag             =   "NoClose"
            Top             =   375
            Width           =   3090
         End
      End
      Begin SuperAv.ShapeButton cmdCekAll 
         Height          =   390
         Left            =   150
         TabIndex        =   105
         Top             =   4425
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         ButtonStyle     =   7
         CaptionAlignment=   3
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Cek All"
         Picture         =   "frmScanVirus.frx":DDBC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdClearAll 
         Height          =   390
         Left            =   1815
         TabIndex        =   106
         Top             =   3900
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         ButtonStyle     =   7
         CaptionAlignment=   3
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Clear All"
         Picture         =   "frmScanVirus.frx":E356
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdFixReg 
         Height          =   390
         Left            =   6360
         TabIndex        =   107
         Top             =   4425
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   688
         ButtonStyle     =   7
         CaptionAlignment=   3
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Fix Registry"
         Picture         =   "frmScanVirus.frx":E8F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "System"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3615
         Left            =   150
         TabIndex        =   90
         Top             =   0
         Width           =   3765
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Task Manager"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   103
            Tag             =   "DisableTaskMgr"
            Top             =   375
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Display Properties"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   7
            Left            =   225
            TabIndex        =   102
            Tag             =   "NoDispCPL"
            Top             =   2010
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Windows Version on Desktop"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   6
            Left            =   225
            TabIndex        =   101
            Tag             =   "PaintDesktopVersion"
            Top             =   1770
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Right-click on Desktop"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   5
            Left            =   225
            TabIndex        =   100
            Tag             =   "NoViewContextMenu"
            Top             =   1530
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Menu Run"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   225
            TabIndex        =   99
            Tag             =   "NoRun"
            Top             =   1290
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Menu Find"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   3
            Left            =   225
            TabIndex        =   98
            Tag             =   "NoFind"
            Top             =   1050
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Folder Options Menu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   225
            TabIndex        =   97
            Tag             =   "NoFolderOptions"
            Top             =   810
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Registry Editor Tools"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   225
            TabIndex        =   96
            Tag             =   "DisableRegistryTools"
            Top             =   570
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Hide And Support"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   12
            Left            =   225
            TabIndex        =   95
            Tag             =   "NoSMHelp"
            Top             =   2250
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Properties My Computer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   13
            Left            =   225
            TabIndex        =   94
            Tag             =   "NoPropertiesMyComputer"
            Top             =   2490
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show File Hidden Operating System "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   14
            Left            =   225
            TabIndex        =   93
            Tag             =   "ShowSuperHidden "
            Top             =   2730
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Hidden Folders And Files "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   15
            Left            =   225
            TabIndex        =   92
            Tag             =   "Hidden "
            Top             =   2970
            Width           =   3090
         End
         Begin VB.CheckBox chkT 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show File Extensions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   16
            Left            =   225
            TabIndex        =   91
            Tag             =   "HideFileExt"
            Top             =   3210
            Width           =   3090
         End
      End
   End
   Begin VB.PictureBox picProcess 
      BackColor       =   &H00FFFFFF&
      Height          =   6090
      Left            =   4200
      Picture         =   "frmScanVirus.frx":EE8A
      ScaleHeight     =   6030
      ScaleWidth      =   10380
      TabIndex        =   49
      Top             =   8160
      Visible         =   0   'False
      Width           =   10440
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Memory Informations"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2865
         Left            =   120
         TabIndex        =   61
         Top             =   4440
         Width           =   7665
         Begin VB.Timer tmrMem 
            Interval        =   1100
            Left            =   3075
            Top             =   2250
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   4050
            TabIndex        =   76
            Top             =   2100
            Width           =   3390
            Begin SuperAv.ProgressBar prgCPU 
               Height          =   315
               Left            =   150
               TabIndex        =   77
               Top             =   225
               Width           =   2640
               _ExtentX        =   4657
               _ExtentY        =   556
               Value           =   0
               Theme           =   8
               TextStyle       =   2
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "prgCPU"
            End
            Begin VB.Label lblCPU 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   2775
               TabIndex        =   78
               Top             =   262
               Width           =   540
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   4050
            TabIndex        =   74
            Top             =   1500
            Width           =   3390
            Begin SuperAv.ProgressBar ProgMemUsed 
               Height          =   315
               Left            =   150
               TabIndex        =   75
               Top             =   225
               Width           =   3090
               _ExtentX        =   5450
               _ExtentY        =   556
               Value           =   0
               Theme           =   8
               TextStyle       =   2
               BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "ProgMemUsed"
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00000000&
            Caption         =   "Kernel Memory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1140
            Left            =   4050
            TabIndex        =   70
            Top             =   300
            Width           =   3390
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   4
               Left            =   150
               TabIndex        =   73
               Top             =   300
               Width           =   3090
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   5
               Left            =   360
               TabIndex        =   72
               Top             =   2400
               Width           =   3090
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   6
               Left            =   150
               TabIndex        =   71
               Top             =   780
               Width           =   3090
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00000000&
            Caption         =   " Virtual Memory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1140
            Left            =   150
            TabIndex        =   66
            Top             =   1575
            Width           =   3390
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   7
               Left            =   150
               TabIndex        =   69
               Top             =   540
               Width           =   3090
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   8
               Left            =   150
               TabIndex        =   68
               Top             =   780
               Width           =   3090
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   9
               Left            =   150
               TabIndex        =   67
               Top             =   300
               Width           =   3090
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00000000&
            Caption         =   "Physical Memory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1140
            Left            =   150
            TabIndex        =   62
            Top             =   300
            Width           =   3390
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   0
               Left            =   150
               TabIndex        =   65
               Top             =   300
               Width           =   3090
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   1
               Left            =   150
               TabIndex        =   64
               Top             =   540
               Width           =   3090
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   2
               Left            =   150
               TabIndex        =   63
               Top             =   780
               Width           =   3090
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1365
         Left            =   150
         TabIndex        =   53
         Top             =   1650
         Width           =   7665
         Begin VB.PictureBox picIconP32 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   480
            Left            =   150
            ScaleHeight     =   69.189
            ScaleMode       =   0  'User
            ScaleWidth      =   298.868
            TabIndex        =   54
            Top             =   225
            Width           =   495
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   825
            TabIndex        =   60
            Top             =   975
            Width           =   615
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   825
            TabIndex        =   59
            Top             =   750
            Width           =   615
         End
         Begin VB.Label lblCompany 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   825
            TabIndex        =   58
            Top             =   450
            Width           =   6390
         End
         Begin VB.Label lblDescription 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   825
            TabIndex        =   57
            Top             =   225
            Width           =   6390
         End
         Begin VB.Label lblFile 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1575
            TabIndex        =   56
            Top             =   750
            Width           =   5790
         End
         Begin VB.Label lblPath 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2640
            TabIndex        =   55
            Top             =   120
            Width           =   5790
         End
      End
      Begin VB.Timer tmrProcessRefresh 
         Interval        =   5000
         Left            =   150
         Top             =   1200
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   840
         Left            =   6825
         TabIndex        =   50
         Top             =   750
         Visible         =   0   'False
         Width           =   990
         Begin VB.PictureBox picIcon 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            ScaleHeight     =   16
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   16
            TabIndex        =   51
            Top             =   450
            Visible         =   0   'False
            Width           =   240
         End
         Begin MSComctlLib.ImageList ImageList3 
            Left            =   375
            Top             =   225
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   4210752
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin VB.Image ImgIcon 
            Height          =   240
            Left            =   75
            Picture         =   "frmScanVirus.frx":10C96
            Top             =   225
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin MSComctlLib.ListView lstView 
         Height          =   1440
         Left            =   120
         TabIndex        =   52
         Top             =   2040
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   2540
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList3"
         SmallIcons      =   "ImageList3"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Process Name"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Directory"
            Object.Width           =   11467
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "User Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Discription"
            Object.Width           =   6880
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Size"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Process ID"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Base P"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Threads"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Attributes"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Priority"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "CRC32"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Mem Usage"
            Object.Width           =   1766
         EndProperty
      End
   End
   Begin VB.PictureBox picStartup 
      BackColor       =   &H00FFFFFF&
      Height          =   6090
      Left            =   5640
      ScaleHeight     =   6030
      ScaleWidth      =   7980
      TabIndex        =   108
      Top             =   8880
      Visible         =   0   'False
      Width           =   8040
      Begin SuperAv.ShapeButton cmdAutorun 
         Height          =   390
         Left            =   6225
         TabIndex        =   115
         Top             =   3705
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         ButtonStyle     =   7
         CaptionAlignment=   3
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Autorun.inf"
         Picture         =   "frmScanVirus.frx":11220
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboStartup 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmScanVirus.frx":1137A
         Left            =   150
         List            =   "frmScanVirus.frx":1139F
         TabIndex        =   114
         Text            =   "All Users Startup"
         Top             =   3750
         Width           =   3390
      End
      Begin VB.TextBox txtPathStartup 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   113
         Top             =   3300
         Width           =   7740
      End
      Begin SuperAv.ShapeButton cmdDelAutorun 
         Height          =   390
         Left            =   6225
         TabIndex        =   116
         Top             =   4170
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         ButtonStyle     =   7
         CaptionAlignment=   3
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Delete"
         Picture         =   "frmScanVirus.frx":11456
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdRefAutorun 
         Height          =   390
         Left            =   6225
         TabIndex        =   117
         Top             =   4635
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         ButtonStyle     =   7
         CaptionAlignment=   3
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Refresh"
         Picture         =   "frmScanVirus.frx":119F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   3090
         Left            =   720
         TabIndex        =   111
         Top             =   360
         Width           =   7740
         Begin VB.ListBox lstStartup 
            BackColor       =   &H8000000D&
            Height          =   2595
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   7290
         End
         Begin VB.ListBox List5 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1380
            ItemData        =   "frmScanVirus.frx":120C2
            Left            =   1200
            List            =   "frmScanVirus.frx":120C4
            TabIndex        =   118
            Top             =   900
            Width           =   660
         End
         Begin VB.ListBox List6 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1380
            Left            =   1950
            TabIndex        =   119
            Top             =   900
            Width           =   660
         End
      End
      Begin VB.Image Image3 
         Height          =   2250
         Left            =   3720
         Picture         =   "frmScanVirus.frx":120C6
         Top             =   3720
         Visible         =   0   'False
         Width           =   2250
      End
   End
   Begin VB.PictureBox picOptions 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6090
      Left            =   4200
      Picture         =   "frmScanVirus.frx":15C8F
      ScaleHeight     =   6090
      ScaleWidth      =   9720
      TabIndex        =   30
      Top             =   2280
      Width           =   9720
      Begin VB.Frame Frame4 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Other Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   2205
         Left            =   3225
         TabIndex        =   43
         Top             =   3570
         Width           =   4590
         Begin VB.CheckBox chkSafeMode 
            BackColor       =   &H00404040&
            Caption         =   "Run In Safe Mode"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   2760
            TabIndex        =   127
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkStartup 
            BackColor       =   &H00404040&
            Caption         =   "Run With Windows (Start Up)"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   1680
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CheckBox chkHideTask 
            BackColor       =   &H00404040&
            Caption         =   "Hide From Task Manager"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   225
            Left            =   225
            MaskColor       =   &H00000000&
            TabIndex        =   120
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chkUnclose 
            BackColor       =   &H00404040&
            Caption         =   "Unclose Send Message"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   225
            Left            =   240
            MaskColor       =   &H00000000&
            TabIndex        =   46
            Top             =   960
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chkOnTop 
            BackColor       =   &H00404040&
            Caption         =   "Always On Top"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   225
            Left            =   225
            MaskColor       =   &H00000000&
            TabIndex        =   45
            Top             =   600
            Width           =   2895
         End
         Begin VB.CheckBox chkTrans 
            BackColor       =   &H00404040&
            Caption         =   "Transparent"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   225
            TabIndex        =   44
            Top             =   240
            Width           =   2595
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Scan Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   2040
         Left            =   3240
         TabIndex        =   39
         Top             =   1350
         Width           =   4590
         Begin VB.CheckBox chkScanRAR 
            BackColor       =   &H00404040&
            Caption         =   "Scan RAR Files"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   225
            TabIndex        =   48
            Top             =   1665
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkSound 
            BackColor       =   &H00404040&
            Caption         =   "Warning Sound"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   225
            TabIndex        =   47
            Top             =   1335
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkCleanAll 
            BackColor       =   &H00404040&
            Caption         =   "Clean All Detected Virus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   240
            Left            =   225
            MaskColor       =   &H00000000&
            TabIndex        =   42
            Top             =   1020
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chkFixRegistry 
            BackColor       =   &H00404040&
            Caption         =   "Fix Infected Registry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   225
            MaskColor       =   &H00000000&
            TabIndex        =   41
            Top             =   705
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chkDisBuff 
            BackColor       =   &H00404040&
            Caption         =   "Disable Buffering"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   225
            TabIndex        =   40
            Top             =   375
            Value           =   1  'Checked
            Width           =   2595
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Filter Extension"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   1140
         Left            =   3225
         TabIndex        =   35
         Top             =   75
         Width           =   4590
         Begin VB.OptionButton optCustomExt 
            BackColor       =   &H00404040&
            Caption         =   "Custom Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   240
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   2265
         End
         Begin VB.OptionButton optAllExt 
            BackColor       =   &H00404040&
            Caption         =   "All Type Extension"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   240
            Left            =   225
            TabIndex        =   37
            Top             =   375
            Value           =   -1  'True
            Width           =   2265
         End
         Begin VB.ComboBox cboExt 
            BackColor       =   &H00404040&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   330
            ItemData        =   "frmScanVirus.frx":188AD
            Left            =   2925
            List            =   "frmScanVirus.frx":188D5
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   645
            Width           =   1470
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Object Database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   5790
         Left            =   360
         TabIndex        =   33
         Top             =   0
         Width           =   2790
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   1950
            Top             =   4950
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmScanVirus.frx":1892B
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lstVirus 
            Height          =   5310
            Left            =   120
            TabIndex        =   34
            Top             =   300
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   9366
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList2"
            ForeColor       =   16777215
            BackColor       =   8421504
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmScanVirus.frx":19605
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Virus Names"
               Object.Width           =   3706
            EndProperty
         End
      End
   End
   Begin VB.PictureBox picScan 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6450
      Left            =   4800
      Picture         =   "frmScanVirus.frx":1991F
      ScaleHeight     =   6450
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   9600
      Begin VB.Timer tmrStatus 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   7560
         Top             =   4680
      End
      Begin VB.TextBox txtDirPath 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   3675
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "frmScanVirus.frx":1B72B
         Top             =   5280
         Width           =   4140
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "STATUS : [ Waiting For Instructions ]"
         Top             =   4800
         Width           =   7665
      End
      Begin VB.Frame fraStatus 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Scanning Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1290
         Left            =   240
         TabIndex        =   5
         Top             =   3375
         Width           =   7665
         Begin VB.Label lblStatus 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   1350
            TabIndex        =   31
            Top             =   375
            Width           =   240
         End
         Begin VB.Label lblTimeValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   ": 00:00:00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   6525
            TabIndex        =   23
            Top             =   375
            Width           =   915
         End
         Begin VB.Label lblStatus 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Elapsed Time"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   4
            Left            =   5100
            TabIndex        =   22
            Top             =   375
            Width           =   1365
         End
         Begin VB.Label lblExt 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   6600
            TabIndex        =   13
            Top             =   615
            Width           =   915
         End
         Begin VB.Label lblStatus 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Extension"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   3
            Left            =   5100
            TabIndex        =   12
            Top             =   615
            Width           =   1365
         End
         Begin VB.Label lblVirusClean 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   ": 0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   1950
            TabIndex        =   11
            Top             =   855
            Width           =   1440
         End
         Begin VB.Label lblVirusDetected 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   ": 0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1950
            TabIndex        =   10
            Top             =   615
            Width           =   1440
         End
         Begin VB.Label lblFileScan 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   ": 0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   1950
            TabIndex        =   9
            Top             =   375
            Width           =   1440
         End
         Begin VB.Label lblStatus 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Object Cleaned"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   225
            TabIndex        =   8
            Top             =   855
            Width           =   1440
         End
         Begin VB.Label lblStatus 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Object Detected"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   225
            TabIndex        =   7
            Top             =   615
            Width           =   1440
         End
         Begin VB.Label lblStatus 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Scanned  File"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   6
            Top             =   375
            Width           =   1140
         End
      End
      Begin VB.PictureBox picDetection 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   150
         ScaleHeight     =   2115
         ScaleWidth      =   7665
         TabIndex        =   1
         Top             =   150
         Width           =   7665
         Begin MSComctlLib.ListView lstDetection 
            Height          =   1890
            Left            =   75
            TabIndex        =   2
            Top             =   75
            Width           =   7440
            _ExtentX        =   13123
            _ExtentY        =   3334
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   255
            BackColor       =   0
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Object Name"
               Object.Width           =   2470
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Directory Location"
               Object.Width           =   8468
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Size ( Byte )"
               Object.Width           =   1940
            EndProperty
         End
      End
      Begin SuperAv.ProgressBar ProgScan 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   3000
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   556
         Value           =   0
         Theme           =   8
         TextStyle       =   4
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Buffering has been disabled by user"
         TextEffect      =   5
      End
      Begin SuperAv.ShapeButton cmdScan 
         Height          =   540
         Left            =   150
         TabIndex        =   16
         Tag             =   "Start Scan"
         ToolTipText     =   "Start Scan"
         Top             =   5250
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   953
         ButtonStyle     =   8
         BackColor       =   33023
         BackColorPressed=   -2147483647
         BackColorHover  =   -2147483647
         BorderColor     =   33023
         BorderColorPressed=   -2147483647
         BorderColorHover=   -2147483647
         Caption         =   ""
         Picture         =   "frmScanVirus.frx":1B74E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdPause 
         Height          =   540
         Left            =   900
         TabIndex        =   17
         Tag             =   "Pause"
         Top             =   5250
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   953
         ButtonStyle     =   8
         BackColor       =   -2147483647
         BackColorPressed=   -2147483647
         BackColorHover  =   -2147483647
         BorderColor     =   10794679
         BorderColorPressed=   -2147483647
         BorderColorHover=   -2147483647
         Caption         =   ""
         Picture         =   "frmScanVirus.frx":1BCE8
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdStop 
         Height          =   540
         Left            =   1650
         TabIndex        =   18
         Tag             =   "Stop Scan"
         Top             =   5250
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   953
         ButtonStyle     =   8
         BackColor       =   33023
         BackColorPressed=   12117984
         BackColorHover  =   33023
         BorderColor     =   10794679
         BorderColorPressed=   12117984
         BorderColorHover=   33023
         Caption         =   ""
         Picture         =   "frmScanVirus.frx":1C282
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SuperAv.ShapeButton cmdBrowse 
         Height          =   540
         Left            =   2700
         TabIndex        =   19
         Tag             =   "Browse Path"
         Top             =   5250
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   953
         ButtonStyle     =   7
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Browse"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPercen 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0 % Completed."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6375
         TabIndex        =   29
         Top             =   2775
         Width           =   1440
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   2550
         X2              =   2550
         Y1              =   5250
         Y2              =   5775
      End
      Begin VB.Label lblFileName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "||-----"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   225
         TabIndex        =   4
         Top             =   2325
         Width           =   7590
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Virus Signature"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9720
      TabIndex        =   133
      Top             =   1095
      Width           =   1440
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Engine Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9720
      TabIndex        =   132
      Top             =   855
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Build "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9720
      TabIndex        =   131
      Top             =   615
      Width           =   1440
   End
   Begin VB.Label lblSystem 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   11370
      TabIndex        =   130
      Top             =   600
      Width           =   1590
   End
   Begin VB.Label lblSystem 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   11370
      TabIndex        =   129
      Top             =   855
      Width           =   1590
   End
   Begin VB.Label lblSystem 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   11370
      TabIndex        =   128
      Top             =   1095
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   8850
      Left            =   -240
      Picture         =   "frmScanVirus.frx":1C81C
      Top             =   -120
      Width           =   17355
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuViewLog 
         Caption         =   "View Log"
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "&Menu"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu mnuShowMe 
         Caption         =   "Show Main Window"
      End
      Begin VB.Menu mnuBatas12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTurn 
         Caption         =   "Turn Off Computer"
         Begin VB.Menu mnuTOC 
            Caption         =   "Turn Off"
            Index           =   1
         End
         Begin VB.Menu mnuTOC 
            Caption         =   "Restart"
            Index           =   2
         End
         Begin VB.Menu mnuTOC 
            Caption         =   "Log Off"
            Index           =   3
         End
      End
      Begin VB.Menu mnuBatas11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEx 
         Caption         =   "Exit"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuScan 
      Caption         =   "&Scan"
      Visible         =   0   'False
      Begin VB.Menu mnuScanProcess 
         Caption         =   "Scan Process And Startup"
      End
      Begin VB.Menu mnuBatas7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScanWindows 
         Caption         =   "Scan Windows"
      End
      Begin VB.Menu mnuScanSystem 
         Caption         =   "Scan System"
      End
      Begin VB.Menu regscan 
         Caption         =   "Scan Network"
      End
      Begin VB.Menu mycomp 
         Caption         =   "Scan USB"
      End
   End
   Begin VB.Menu mnuPrio 
      Caption         =   "&Priority"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuPriority 
         Caption         =   "&1. Realtime Priority"
         Index           =   1
      End
      Begin VB.Menu mnuPriority 
         Caption         =   "&2. High Priority"
         Index           =   2
      End
      Begin VB.Menu mnuPriority 
         Caption         =   "&3. Normal Priority"
         Index           =   3
      End
      Begin VB.Menu mnuPriority 
         Caption         =   "&4. Idle Priority"
         Index           =   4
      End
   End
   Begin VB.Menu mnuVirus 
      Caption         =   "&Virus"
      Visible         =   0   'False
      Begin VB.Menu mnuViri 
         Caption         =   "Clean/Delete"
         Index           =   1
      End
      Begin VB.Menu mnuViri 
         Caption         =   "Quarantine"
         Index           =   2
      End
      Begin VB.Menu mnuBatas1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViriSelect 
         Caption         =   "Select All"
         Index           =   1
      End
      Begin VB.Menu mnuViriSelect 
         Caption         =   "Unselect"
         Index           =   2
      End
      Begin VB.Menu mnuBatas2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCleanAllViri 
         Caption         =   "Clean All Selected"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocation 
         Caption         =   "File Location"
      End
   End
   Begin VB.Menu mnuT 
      Caption         =   "&Processes / Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuNewProcess 
         Caption         =   "New Process"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh Process"
      End
      Begin VB.Menu mnBatas3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCur 
         Caption         =   "Threads"
         Begin VB.Menu mnuThread 
            Caption         =   "Resume Process"
            Index           =   1
         End
         Begin VB.Menu mnuThread 
            Caption         =   "Suspend Process"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEndProcess 
         Caption         =   "End Process"
      End
      Begin VB.Menu mnuBatas4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetPrio 
         Caption         =   "Set Process Priority"
         Begin VB.Menu mnuSetPriority 
            Caption         =   "&1. Realtime Priority"
            Index           =   1
         End
         Begin VB.Menu mnuSetPriority 
            Caption         =   "&2. High Priority"
            Index           =   2
         End
         Begin VB.Menu mnuSetPriority 
            Caption         =   "&3. Normal Priority"
            Index           =   3
         End
         Begin VB.Menu mnuSetPriority 
            Caption         =   "&4. Idle Priority"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "Show File Informations"
      End
      Begin VB.Menu mnuBatas5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindFile 
         Caption         =   "Find File Location"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuQua 
      Caption         =   "&Quarantine"
      Visible         =   0   'False
      Begin VB.Menu mnuClean 
         Caption         =   "Clean All Object "
         Index           =   1
      End
      Begin VB.Menu mnuClean 
         Caption         =   "Clean Selected"
         Index           =   2
      End
   End
   Begin VB.Menu extra 
      Caption         =   "&Extra Tools"
      Visible         =   0   'False
      Begin VB.Menu script 
         Caption         =   "Script Cleaner"
      End
      Begin VB.Menu sma 
         Caption         =   "Smad Av"
      End
      Begin VB.Menu usbb 
         Caption         =   "USB DisInfector"
      End
      Begin VB.Menu junk 
         Caption         =   "Clean Junk Files"
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Windows"
      Visible         =   0   'False
      Begin VB.Menu mnuWindows 
         Caption         =   "Console Windows"
         Index           =   1
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "System Configurations"
         Index           =   2
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "Task Manager"
         Index           =   3
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "Registry Editor"
         Index           =   4
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "System Restore"
         Index           =   5
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "Clean Menager"
         Index           =   6
      End
      Begin VB.Menu mnuBatas6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExplorer 
         Caption         =   "Windows Explorer"
      End
      Begin VB.Menu mnuControlPanel 
         Caption         =   "Control Panel"
      End
   End
   Begin VB.Menu upd 
      Caption         =   "&Online Update"
      Visible         =   0   'False
      Begin VB.Menu how 
         Caption         =   "How To Update"
      End
      Begin VB.Menu updd 
         Caption         =   "Download Update File"
      End
   End
   Begin VB.Menu mnuABout 
      Caption         =   "&About"
      Visible         =   0   'False
      Begin VB.Menu mnucheckweb 
         Caption         =   "Check WebSite"
      End
      Begin VB.Menu mnuGP 
         Caption         =   "Tips Tricks"
      End
      Begin VB.Menu mnuGR 
         Caption         =   "About Super Av"
      End
   End
End
Attribute VB_Name = "frmScanVirus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Av Super Protector - A N T I V I R U S



Private m_hMod As Long

Dim Hours As String
Dim Minutes As String
Dim Seconds As String
Dim MilliSec As String
Dim ElapsedMilliSec As Long
Dim TotalElapsedMilliSec As Long
Dim StartTickCount As Long
Dim Seal As New clsHuffman
Dim WhereMine As String
Dim lAlpha As Integer

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub cboStartup_Click()
    GetAllRun
    cmdDelAutorun.Enabled = False
    txtPathStartup.Text = ""
End Sub

Private Sub chkCleanAll_Click()
    If chkCleanAll.value = 1 Then
        DeleteAll = True
    Else
        DeleteAll = False
    End If
End Sub

Private Sub chkFixRegistry_Click()
    If chkFixRegistry.value = 1 Then
        RegistryFix = True
    Else
        RegistryFix = False
    End If
End Sub

Private Sub chkHideTask_Click()
    If chkHideTask.value = 1 Then
        lTuan = GetWindow(Me.hWnd, GW_OWNER)
        ShowWindow lTuan, SW_HIDE
    Else
        App.Title = " "
        ShowWindow lTuan, SW_SHOW
    End If
End Sub

Private Sub chkOnTop_Click()
    If chkOnTop.value = 1 Then
        AlwaysOnTop Me.hWnd, True
    Else
        AlwaysOnTop Me.hWnd, False
    End If
End Sub







Private Sub chkT_Click(Index As Integer)
    On Error Resume Next
    
    If cekLoad = True Then
        CekSetting = True
        cmdApplyTweak.Enabled = True
        cmdApplyTweak.Caption = "Apply"
    End If
End Sub

Private Sub chkTrans_Click()
    If chkTrans.value = 1 Then
        SetTrans Me, 125
    Else
        chkTrans.value = 0
        SetTrans Me, 255
    End If
End Sub

Private Sub chkUnclose_Click()
    If chkUnclose.value = 1 Then Exit Sub
End Sub

Private Sub cleen_Click(Index As Integer)
'ShellExecute Me.hwnd, vbNullString, "rcleanmgr.exe", vbNullString, "C:\", 1
End Sub

Private Sub cmdApplyTweak_Click()
    SaveApp
    cmdApplyTweak.Enabled = False
    cmdApplyTweak.Caption = "No Changes"
    LockWindowUpdate (GetDesktopWindow())
    ForceCacheRefresh
    LockWindowUpdate (0)
End Sub

Private Sub cmdAutorun_Click()
    If MsgBox("Are you sure to delete autorun.inf in all drives ?", vbYesNo + vbQuestion, APP_PROGRAM & " /Delete Autorun") = vbYes Then
         ClearAuto
         Call MsgBox("All autorun.inf was deleted !", vbYesNo + vbInformation, APP_PROGRAM)
    End If
End Sub

Private Sub cmdBrowse_Click()
    Where = BrowseForFolder(Me.hWnd, "Select Drive or Directory to scan :")
    If Len(Where) > 0 Then
        txtDirPath = Where
    End If
End Sub

Private Sub cmdBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdBrowse.Tag
End Sub

Private Sub cmdCekAll_Click()
    Dim I As Integer
    
    On Error Resume Next
    
    With chkT
        For I = 0 To .count
            .Item(I).value = 1
        Next I
    End With
End Sub

Private Sub cmdClearAll_Click()
    Dim I As Integer
    
    On Error Resume Next
    
    With chkT
        For I = 0 To .count
            .Item(I).value = 0
        Next I
    End With
End Sub

Private Sub cmdDelAutorun_Click()
    If MsgBox("Are you sure..?", vbQuestion + vbYesNo, "/Delete Startup") = vbYes Then
        ClearAutorun
        cmdDelAutorun.Enabled = False
        txtPathStartup.Text = ""
    End If
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, APP_PROGRAM) = vbYes Then
       ' MsgBox "Thanks For Using " & APP_PROGRAM, vbSystemModal + vbInformation, APP_PROGRAM
        tmrFadeout.Enabled = True
        SystrayOff Me
        frmWait.Show vbModal
        FadeOut Me
        End
    End If
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdExit.Tag
End Sub

Private Sub cmdFixReg_Click()
    If tmrFix.Enabled = False Then
        If MsgBox("Are you sure want to fix the registry ?", vbExclamation + vbYesNo, "- Fix Registry") = vbYes Then
            tmrFix.Enabled = True
        End If
    Else
        tmrFix.Enabled = False
        txtStatus.Text = "Your Computer Is in Good Conditions"
    End If
End Sub

Private Sub cmdMenuAbout_Click()
    frmQuarantine.Show
    Me.Enabled = False
End Sub

Private Sub cmdMenuAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdMenuAbout.Tag
End Sub

Private Sub cmdMenuOptions_Click()
ShapeButton1.Visible = False
ShapeButton2.Visible = False
ShapeButton3.Visible = False

    M_Scan (False): M_Options (True): M_Process (False): M_Tweak (False): M_Startup (False)
End Sub

Private Sub cmdMenuOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(1).Text = cmdMenuOptions.Tag
End Sub

Private Sub cmdMenuProcess_Click()
    M_Scan (False): M_Options (False): M_Process (True): M_Tweak (False): M_Startup (False)
End Sub

Private Sub cmdMenuProcess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdMenuProcess.Tag
End Sub

Private Sub cmdMenuScan_Click()
picScan.Visible = True
ShapeButton1.Visible = False
ShapeButton2.Visible = False
ShapeButton3.Visible = False




    M_Scan (True): M_Options (False): M_Process (False): M_Tweak (False): M_Startup (False)
End Sub

Private Sub cmdMenuScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdMenuScan.Tag
End Sub

Private Sub cmdMenuStartup_Click()
    M_Scan (False): M_Options (False): M_Process (False): M_Tweak (False): M_Startup (True)
End Sub

Private Sub cmdMenuStartup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdMenuStartup.Tag
End Sub

Private Sub cmdMenuTweakReg_Click()
    M_Scan (False): M_Options (False): M_Process (False): M_Tweak (True): M_Startup (False)
End Sub

Private Sub cmdMenuTweakReg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdMenuTweakReg.Tag
End Sub

Private Sub cmdPause_Click()
    StopScan = False
    PauseScan = True
    tmrStatus.Enabled = False
    cmdPause.Enabled = False
    cmdScan.Enabled = True
End Sub

Private Sub cmdPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdPause.Tag
End Sub

Private Sub cmdQuarantine_Click()
    frmQuarantine.Show
    Me.Enabled = False
End Sub

Private Sub cmdQuarantine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdQuarantine.Tag
End Sub

Private Sub cmdRefAutorun_Click()
    GetAllRun
    cmdDelAutorun.Enabled = False
    txtPathStartup.Text = ""
End Sub

Private Sub cmdScan_Click()
    If Mid$(txtDirPath, 2, 1) <> ":" Then
        MsgBox "Path not found!", vbExclamation, APP_PROGRAM
        Exit Sub
    End If
    If StopButton = True Then
        PauseScan = False
        cmdPause.Enabled = True
        cmdScan.Enabled = False
        TotalElapsedMilliSec = TotalElapsedMilliSec + (GetTickCount() - StartTickCount)
        TotalElapsedMilliSec = 0
        tmrStatus.Enabled = True
    Else
        LogFile "Scanning in     " & txtDirPath
        lstDetection.ListItems.Clear
        lblFileScan.Caption = ": 0"
        lblVirusDetected.Caption = ": 0"
        lblVirusClean.Caption = ": 0"
        StartTickCount = GetTickCount()
        tmrStatus.Enabled = False
        CheckItem
        ext = cboExt.Text
        cmdPause.Enabled = True
        cmdStop.Enabled = True
        ProcedureScan
    End If
End Sub

Private Sub cmdScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdScan.Tag
End Sub

Private Sub cmdSetAttrib_Click()
    frmSetAttrib.Show
  '  Me.Enabled = False
  Me.Hide
  
End Sub

Private Sub cmdSignature_Click()
    frmSignature.Show
    Me.Enabled = False
End Sub

Private Sub cmdSignature_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdSignature.Tag
End Sub

Private Sub cmdStop_Click()
    If tmrStatus = True Then
        StopScan = True
        cmdScan.Enabled = True
        cmdPause.Enabled = True
        cmdStop.Enabled = True
        cmdBrowse.Enabled = True
    Else
        PauseScan = True
        cmdScan.Enabled = True
        cmdPause.Enabled = True
        cmdStop.Enabled = True
        cmdBrowse.Enabled = True
        Buka
    End If
End Sub

Private Sub cmdStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdStop.Tag
End Sub

Private Sub cmdSystray_Click()
    Dim sTitle As String
    Dim sMessage As String
    
    sTitle = APP_PROGRAM & " V3"
    sMessage = "Copyright © Ermal Gjermeni Softwares 2010" & vbCrLf & _
                vbCrLf & _
                "* App Version : " & APP_VERSION & vbCrLf & _
                "* Current Build : " & CURRENT_BUILD & vbCrLf & _
                "* Processes : " & PROCESSESES & vbCrLf & _
                "* Scan Engine : " & ENGINE_VERSION & vbCrLf & _
                "* Tweak Registry : " & TWEAK_REG & vbCrLf & _
                "* Virus Signature : " & VirusName.count & " Viruses" & vbCrLf & _
                vbCrLf & _
                "Click To Close"

    SystrayOn Me, sTitle
    PopupBalloon Me, sMessage, sTitle, NIIF_INFO
    frmStartup.Visible = False
    With Me
        .Hide
        .Enabled = False
        .tmrFadeout = False
    End With
End Sub

Private Sub cmdSystray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(1).Text = cmdSystray.Tag
End Sub

Private Sub Form_Initialize()
    m_hMod = LoadLibrary("shell32.dll")
    InitCommonControls
End Sub

Private Sub Form_Load()
    Me.Visible = False
    If App.PrevInstance = True Then
        FreeLibrary m_hMod
        MsgBox "Program Already Runing on Your Machine"
        End
        Unload Me
        Close
        
    End If
    
    PauseScan = False
    StopButton = False
    StopScan = False
    LoadExternalDatabase True
    LoadVirusDatabase
    LoadBinaryIconCompare
    cekLoad = False
    CekSetting = False
    GetApp
    cekLoad = True
    GetAllRun
    lAlpha = 255
    
    ' Self protection ----------------------------------
    Me.Caption = "Av Super Protector"
    chkHideTask_Click
    
    chkOnTop_Click
    M_Scan (False): M_Options (False): M_Process (False): M_Tweak (False): M_Startup (False)
    'FadeIn Me
    cboExt.Text = "*.*"
    strUserCom = GetUserCom
    MemoryInfo lblInfo(0), lblInfo(1), lblInfo(2), Frame15, lblInfo(4), lblInfo(5), lblInfo(6), lblInfo(7), lblInfo(8), lblInfo(9), ProgMemUsed, Me.StatusBar1
    GetCPUInfo lblCPU, prgCPU, Me.StatusBar1
    ViewProcess
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Action As Long

    If Me.ScaleMode = vbPixels Then
        Action = X
      Else
        Action = X / Screen.TwipsPerPixelX
    End If

    Select Case Action
      Case WM_RBUTTONUP
        PopupMenu Menu
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' If chkUnclose.value = 0 Then Exit Sub
   ' Cancel = 1
End Sub

Private Sub Form_Terminate()
If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, APP_PROGRAM) = vbYes Then
       ' MsgBox "Thanks For Using " & APP_PROGRAM, vbSystemModal + vbInformation, APP_PROGRAM
        tmrFadeout.Enabled = True
        SystrayOff Me
        frmWait.Show vbModal
        FadeOut Me
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, APP_PROGRAM) = vbYes Then
       ' MsgBox "Thanks For Using " & APP_PROGRAM, vbSystemModal + vbInformation, APP_PROGRAM
        tmrFadeout.Enabled = True
        SystrayOff Me
        frmWait.Show vbModal
        FadeOut Me
        End
    End If
End Sub

Private Sub how_Click()
' tutorial
MsgBox "   To update this program is very easy." & vbCrLf & _
       "1. Download the .TCM.rar (Virus Data Base)& Unrar it file" & vbCrLf & _
       "2. EXIT Av Super Protector." & vbCrLf & _
       "3. Put TCM file where the main program is installed." & vbCrLf & _
       "    Usually C:\Program Files\Ermal Gjermeni Softwares\Super Antivirus\" & vbCrLf & _
       "4. Replace old file with ne new one..." & vbCrLf & _
       "5. Run Super Antivirus again"
End Sub

Private Sub junk_Click()
ShellExecute Me.hWnd, vbNullString, nPath(App.path) & "\Clean temp.bat", vbNullString, "C:\", 1
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub lstDetection_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If lstDetection.ListItems.count > 0 Then
            mnuFileLocation.Tag = lstDetection.SelectedItem.SubItems(1)
            PopupMenu mnuVirus
        End If
    End If
End Sub

Private Sub lstStartup_Click()

    On Error Resume Next

    List5.Selected(lstStartup.ListIndex) = True
    List6.Selected(lstStartup.ListIndex) = True
    txtPathStartup.Text = List5.Text
    If txtPathStartup.Text <> "" Then cmdDelAutorun.Enabled = True
    
End Sub

Private Sub lstView_Click()
    Dim strFile As String, uProcess As PROCESSENTRY32
    Dim hVer As VERHEADER
    Dim fso As New FileSystemObject, FileInfo As file
    Dim strF As String
    
    picIconP32.Cls
    strFile = lstView.SelectedItem.SubItems(1)
    
    If strF <> strFile Then
        On Error GoTo SalahProses
        Set FileInfo = fso.GetFile(strFile)
        GetVerHeader strFile, hVer
    
        Label8.Caption = "File"
        Label7.Caption = "Folder"
    
        lblDescription.Caption = hVer.FileDescription
        lblCompany.Caption = hVer.CompanyName
        lblFile.Caption = ": " & FileInfo.ShortName ' GetFileName(strFile)
        lblPath.Caption = ": " & FileInfo.ParentFolder ' GetFilePath(strFile)
        RetrieveIcon strFile, picIconP32, ricnLarge
        Exit Sub
    End If
    
SalahProses:
        MsgBox Err.Description & " " & " " & _
                "or file has been deleted.", vbExclamation, "Warning"
        mnuRefresh_Click
End Sub

Private Sub lstView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    lstView.Sorted = True
    
    lstView.SortKey = ColumnHeader.Index - 1
    If lstView.SortOrder = lvwDescending Then
       lstView.SortOrder = lvwAscending
    Else
       lstView.SortOrder = lvwDescending
    End If

End Sub

Private Sub lstView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 2 Then
      If lstView.ListItems.count > 0 Then
         
         mnuFindFile.Caption = "Find File Location..."
         mnuFindFile.Tag = lstView.SelectedItem.SubItems(1)
         
            mnuSetPriority(1).Checked = False
            mnuSetPriority(2).Checked = False
            mnuSetPriority(3).Checked = False
            mnuSetPriority(4).Checked = False
            
            Dim priHwnd  As Long
            priHwnd = GetPriority(CLng(lstView.SelectedItem.Tag))
            Select Case priHwnd
                   Case REALTIME_PRIORITY_CLASS
                        mnuSetPriority(1).Checked = True
                   Case HIGH_PRIORITY_CLASS
                        mnuSetPriority(2).Checked = True
                   Case NORMAL_PRIORITY_CLASS
                        mnuSetPriority(3).Checked = True
                   Case IDLE_PRIORITY_CLASS
                        mnuSetPriority(4).Checked = True
            End Select
         PopupMenu mnuT
      End If
    End If
End Sub

Private Sub mnucheckweb_Click()
ShellExecute 0, "open", "http://superav.sourceforge.net", vbNullString, vbNullString, 1    ' here is the style
End Sub

Private Sub mnuCleanAllViri_Click()
    Dim lValue As Long
    
    lValue = CheckVirus
    
    If lValue > 0 Then
        If MsgBox("Are you sure want to clean selected object(s) ?", vbYesNo + vbQuestion, APP_PROGRAM) = vbYes Then
            txtStatus.Text = "STATUS : Cleaning All Object(s)."
            CleanVirus
        End If
    End If
End Sub

Private Sub mnuControlPanel_Click()
    ShellExecute Me.hWnd, vbNullString, "control.exe", vbNullString, "C:\", 1
End Sub

Private Sub mnuEndProcess_Click()
    Dim I As Integer
    Dim Pesan As String, strFile As String
    Dim fso As New FileSystemObject, FileName As file
    
    strFile = lstView.SelectedItem.SubItems(1)
    Set FileName = fso.GetFile(strFile)
    
    Pesan = "WARNING: Terminating a process can cause undesired" & vbCrLf & _
            "results including loss of data and system instability. The" & vbCrLf & _
            "process will not be given the chance to save its state or" & vbCrLf & _
            "data before it is terminated." & vbCrLf & vbCrLf & _
            "Are you sure you want to terminate process" & " " & FileName.ShortName
            If MsgBox(Pesan, vbYesNo + 48, APP_PROGRAM & " /Confirm" & Chr(0)) = vbYes Then
               Dim H As Long
                   H = lstView.SelectedItem.Index
                    For I = 1 To lstView.ListItems.count
                      If lstView.ListItems(I).Selected Then
                        Call KillProcessById(CLng(lstView.ListItems(I).Tag))
                        Sleep 100
                      End If
                    Next I
            End If
    ViewProcess
End Sub

Private Sub mnuEx_Click()
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, APP_PROGRAM) = vbYes Then
        MsgBox "Thanks for using " & APP_PROGRAM, vbSystemModal + vbInformation, APP_PROGRAM
        tmrFadeout.Enabled = True
        frmWait.Show vbModal
        FadeOut Me
        End
    End If
End Sub

Private Sub mnuExplorer_Click()
    ShellExecute Me.hWnd, vbNullString, "explorer.exe", vbNullString, "C:\", 1
End Sub

Private Sub mnufileinfo_Click()
    frmWait.Show vbModal
    frmModule.Show vbModal
End Sub

Private Sub mnuFileLocation_Click()
    On Error Resume Next
    If Trim(mnuFileLocation.Tag) <> "" Then
        Shell "explorer.exe /select," & mnuFileLocation.Tag, 1
    End If
End Sub

Private Sub mnuFindFile_Click()

    frmWait.Show vbModal
    On Error Resume Next
    
    If Trim(mnuFindFile.Tag) <> "" Then
        Shell "explorer.exe /select, " & mnuFindFile.Tag, 1
    End If

End Sub

Private Sub mnuGP_Click()
  
        frmGPL.Show
End Sub

Private Sub mnuGR_Click()
    Me.Enabled = False
    frmAbout.Show
End Sub

Private Sub mnuNewProcess_Click()
    Dim sTitle As String, sPrompt As String
    
    sTitle = APP_PROGRAM & " " & "/New Process "
    sPrompt = "Type the name of a program, folder, document, or Internet resource."
                
    If IsWinNT Then
        SHRunDialog Me.hWnd, 0, 0, StrConv(sTitle, vbUnicode), StrConv(sPrompt, vbUnicode), 0
    Else
        SHRunDialog Me.hWnd, 0, 0, sTitle, sPrompt, 0
    End If
End Sub
Private Sub mnuT_Click()
M_Scan (False): M_Options (False): M_Process (True): M_Tweak (False): M_Startup (False)
End Sub
    

Private Sub mnuPriority_Click(Index As Integer)
    On Error Resume Next
    
    Dim priHwnd  As Long
    Dim insel As Long
    
    Select Case Index
           Case 1
                insel = REALTIME_PRIORITY_CLASS
           Case 2
                insel = HIGH_PRIORITY_CLASS
           Case 3
                insel = NORMAL_PRIORITY_CLASS
           Case 4
                insel = IDLE_PRIORITY_CLASS
    End Select
    
    Dim I As Integer
    If insel <> 0 Then
        For I = 1 To lstView.ListItems.count
            If lstView.ListItems(I).Selected Then
                priHwnd = OpenProcess(PROCESS_SET_INFORMATION, False, CLng(lstView.ListItems(I).Tag))
                SetPriorityClass priHwnd, insel
                CloseHandle priHwnd
            End If
        Next I
    End If
End Sub

Private Sub mnuProperties_Click()
    On Error Resume Next
    
    Dim I As Integer
    
    For I = 1 To lstView.ListItems.count
      If lstView.ListItems(I).Selected Then
         ShowProps lstView.ListItems(I).SubItems(1), Me.hWnd
      End If
    Next I
End Sub

Private Sub mnuRefresh_Click()
    ClearLabel
    ViewProcess
End Sub

Private Sub mnuScanProcess_Click()
 '   lstDetection.ListItems.Clear
 '   FadeIn frmStartup
 '   frmStartup.Show
  '  frmStartup.lblLoad.Caption = "Scanning Processes And Startup"
 '   LogFile "Scanning in     " & GetSpecialfolder(CSIDL_STARTUP)
   ' ScanProcess False
    Where = GetSpecialfolder(CSIDL_STARTUP)
    LogFile "Scanning in     " & Where
    lstDetection.ListItems.Clear
    lblFileScan.Caption = ": 0"
    lblVirusDetected.Caption = ": 0"
    lblVirusClean.Caption = ": 0"
    StartTickCount = GetTickCount()
    tmrStatus.Enabled = False
    CheckItem
    ext = cboExt.Text
    cmdPause.Enabled = False
    cmdStop.Enabled = True
    ProcedureScan
    
End Sub

Private Sub mnuScanSystem_Click()
    Where = GetSpecialfolder(CSIDL_SYSTEM)
    LogFile "Scanning in     " & Where
    lstDetection.ListItems.Clear
    lblFileScan.Caption = ": 0"
    lblVirusDetected.Caption = ": 0"
    lblVirusClean.Caption = ": 0"
    StartTickCount = GetTickCount()
    tmrStatus.Enabled = False
    CheckItem
    ext = cboExt.Text
    cmdPause.Enabled = False
    cmdStop.Enabled = True
    ProcedureScan
End Sub

Private Sub mnuScanWindows_Click()
    Where = GetSpecialfolder(CSIDL_WINDOWS)
    LogFile "Scanning in     " & Where
    lstDetection.ListItems.Clear
    lblFileScan.Caption = ": 0"
    lblVirusDetected.Caption = ": 0"
    lblVirusClean.Caption = ": 0"
    StartTickCount = GetTickCount()
    tmrStatus.Enabled = False
    CheckItem
    ext = cboExt.Text
    cmdPause.Enabled = False
    cmdStop.Enabled = True
    ProcedureScan
End Sub

Private Sub mnuSetPriority_Click(Index As Integer)
    On Error Resume Next
    
    Dim priHwnd  As Long
    Dim insel As Long
    
    Select Case Index
           Case 1
                insel = REALTIME_PRIORITY_CLASS
           Case 2
                insel = HIGH_PRIORITY_CLASS
           Case 3
                insel = NORMAL_PRIORITY_CLASS
           Case 4
                insel = IDLE_PRIORITY_CLASS
    End Select
    
    Dim I As Integer
    If insel <> 0 Then
        For I = 1 To lstView.ListItems.count
          If lstView.ListItems(I).Selected Then
             priHwnd = OpenProcess(PROCESS_SET_INFORMATION, False, CLng(lstView.ListItems(I).Tag))
             SetPriorityClass priHwnd, insel
             CloseHandle priHwnd
          End If
        Next I
    ViewProcess
    End If
End Sub

Private Sub mnuShowMe_Click()
    SystrayOff Me
    frmStartup.Visible = False
    With Me
        .Show
        .Enabled = True
        .tmrFadeout = False
    End With
End Sub

Private Sub mnuThread_Click(Index As Integer)
    Select Case Index
        Case 1: SetSuspendResumeThread lstView, 5, False
        Case 2: SetSuspendResumeThread lstView, 5, True
    End Select
End Sub

Private Sub mnuTOC_Click(Index As Integer)
Select Case Index
    Case 1: SHShutDownDialog 0
    Case 2: SHShutDownDialog 0
    Case 3: LogOffNT True
End Select
End Sub

Private Sub mnuViewLog_Click()
    On Error Resume Next
    Dim ss As String
    If App.path & "\Log\" & "AvLog" & ".txt" Then
        ss = App.path & "\Log\" & "AvLog" & ".txt"
        Dim I As Long
        I = ShellExecute(hWnd, "open", "notepad", ss, "", SW_SHOWNORMAL)
    End If
End Sub

Private Sub mnuViri_Click(Index As Integer)
    Dim lValue As Long
    
    lValue = CheckVirus
    Select Case Index
        Case 1
            If lValue > 0 Then
                txtStatus.ForeColor = &H80000008
                txtStatus.Text = "STATUS : Cleaning Files."
                CleanVirus
            End If
        Case 2: Quarantine
    End Select
End Sub

Private Sub mnuViriSelect_Click(Index As Integer)
    Select Case Index
        Case 1: SelectAll
        Case 2: Unselect
    End Select
End Sub

Private Sub mnuWindows_Click(Index As Integer)
    On Error Resume Next
    
    Select Case Index
        Case 1
            ShellExecute Me.hWnd, vbNullString, "cmd.exe", vbNullString, "C:\", 1
        Case 2
            ShellExecute Me.hWnd, vbNullString, "msconfig.exe", vbNullString, "C:\", 1
        Case 3
            ShellExecute Me.hWnd, vbNullString, "taskmgr.exe", vbNullString, "C:\", 1
        Case 4
            ShellExecute Me.hWnd, vbNullString, "regedit.exe", vbNullString, "C:\", 1
        Case 5
            ShellExecute Me.hWnd, vbNullString, MyWindowSys & "restore\rstrui.exe", vbNullString, "C:\", 1
        Case 6
            ShellExecute Me.hWnd, vbNullString, "cleanmgr.exe", vbNullString, "C:\", 1
    
    End Select
End Sub

Private Sub mycomp_Click()
'    Where = BrowseForFolder(Me.hwnd, "Select Drive or Directory to scan :")
'    If Len(Where) > 0 Then
'        txtDirPath = Where
 '   End If
 
 Where = GetSpecialfolder(CSIDL_DRIVES)
    LogFile "Scanning in     " & Where
    lstDetection.ListItems.Clear
    lblFileScan.Caption = ": 0"
    lblVirusDetected.Caption = ": 0"
    lblVirusClean.Caption = ": 0"
    StartTickCount = GetTickCount()
    tmrStatus.Enabled = False
    CheckItem
    ext = cboExt.Text
    cmdPause.Enabled = False
    cmdStop.Enabled = True
    ProcedureScan
        
End Sub

Private Sub optAllExt_Click()
    cboExt.Text = "*.*"
    cboExt.Enabled = False
End Sub

Private Sub optCustomExt_Click()
    cboExt.Enabled = True
End Sub

Private Sub regscan_Click()
 Where = GetSpecialfolder(CSIDL_NETWORK)
 LogFile "Scanning in     " & Where
    lstDetection.ListItems.Clear
    lblFileScan.Caption = ": 0"
    lblVirusDetected.Caption = ": 0"
    lblVirusClean.Caption = ": 0"
    StartTickCount = GetTickCount()
    tmrStatus.Enabled = False
    CheckItem
    ext = cboExt.Text
    cmdPause.Enabled = False
    cmdStop.Enabled = True
    ProcedureScan
End Sub

Private Sub script_Click()
ShellExecute Me.hWnd, vbNullString, nPath(App.path) & "\killer.vbs", vbNullString, "C:\", 1
End Sub

Private Sub ShapeButton2_Click()
picScan.Visible = True
ShapeButton2.Visible = False
ShapeButton1.Visible = False
ShapeButton3.Visible = False


Where = GetSpecialfolder(CSIDL_DRIVES)
    LogFile "Scanning in     " & Where
    lstDetection.ListItems.Clear
    lblFileScan.Caption = ": 0"
    lblVirusDetected.Caption = ": 0"
    lblVirusClean.Caption = ": 0"
    StartTickCount = GetTickCount()
    tmrStatus.Enabled = False
    CheckItem
    ext = cboExt.Text
    cmdPause.Enabled = False
    cmdStop.Enabled = True
    ProcedureScan
End Sub

Private Sub ShapeButton1_Click()
picScan.Visible = True
ShapeButton1.Visible = False
ShapeButton2.Visible = False
ShapeButton3.Visible = False



 Where = GetSpecialfolder(CSIDL_WINDOWS)
    LogFile "Scanning in     " & Where
    lstDetection.ListItems.Clear
    lblFileScan.Caption = ": 0"
    lblVirusDetected.Caption = ": 0"
    lblVirusClean.Caption = ": 0"
    StartTickCount = GetTickCount()
    tmrStatus.Enabled = False
    CheckItem
    ext = cboExt.Text
    cmdPause.Enabled = False
    cmdStop.Enabled = True
    ProcedureScan
End Sub

Private Sub ShapeButton3_Click()
picScan.Visible = True
ShapeButton3.Visible = False
ShapeButton1.Visible = False
ShapeButton2.Visible = False


 Where = GetSpecialfolder(CSIDL_SYSTEM)
    LogFile "Scanning in     " & Where
    lstDetection.ListItems.Clear
    lblFileScan.Caption = ": 0"
    lblVirusDetected.Caption = ": 0"
    lblVirusClean.Caption = ": 0"
    StartTickCount = GetTickCount()
    tmrStatus.Enabled = False
    CheckItem
    ext = cboExt.Text
    cmdPause.Enabled = False
    cmdStop.Enabled = True
    ProcedureScan
End Sub

Private Sub ShapeButton4_Click()
picScan.Visible = False
picOptions.Visible = False




ShapeButton1.Visible = True
ShapeButton2.Visible = True
ShapeButton3.Visible = True



End Sub

Private Sub ShapeButton5_Click()
frmAbout.Show
End Sub

Private Sub sma_Click()
Reg.Show
ShellExecute Me.hWnd, vbNullString, nPath(App.path) & "\Smadav.exe", vbNullString, "C:\", 1
End Sub

Private Sub Text1_Change()

End Sub

Private Sub tmrFadeout_Timer()
    If lAlpha > 0 Then
        DoEvents
        lAlpha = lAlpha - 5
        MakeTransparent Me.hWnd, lAlpha
    Else
        lAlpha = 0
        Me.Hide
        tmrFadeout.Enabled = False
        End
    End If
End Sub

Private Sub tmrFix_Timer()
    If progFixReg.value < 30 Then
        progFixReg.value = progFixReg.value + 1
    ElseIf progFixReg.value < 60 Then
        progFixReg.value = 75
    ElseIf progFixReg.value < 75 Then
        progFixReg.value = 85
    ElseIf progFixReg.value < 85 Then
        progFixReg.value = 95
    ElseIf progFixReg.value = 95 Then
        progFixReg.value = 100
        FixRegistry
        tmrFix.Enabled = False
       ' MsgBox "Your Computer Is in Good Conditions "
        
        progFixReg.value = 0
    End If
End Sub

Private Sub tmrMem_Timer()
    DoEvents
    
    UpdateValues lblCPU, prgCPU, Me.StatusBar1
    MemoryInfo lblInfo(0), lblInfo(1), lblInfo(2), Frame15, lblInfo(4), lblInfo(5), _
                lblInfo(6), lblInfo(7), lblInfo(8), lblInfo(9), ProgMemUsed, Me.StatusBar1
End Sub

Private Sub tmrStatus_Timer()
    If txtStatus.ForeColor = &HFF0000 Then
        txtStatus.ForeColor = &HFF&
    Else
        txtStatus.ForeColor = &HFF0000
    End If
    
    ElapsedMilliSec = (GetTickCount() - StartTickCount) + TotalElapsedMilliSec
    Hours = Format((ElapsedMilliSec \ 3600000), "00")
    Minutes = Format((ElapsedMilliSec \ 60000) Mod 60, "00")
    Seconds = Format((ElapsedMilliSec \ 1000) Mod 60, "00")
    MilliSec = Format((ElapsedMilliSec Mod 1000) \ 10, "00")
    lblTimeValue = ": " & Hours & ":" & Minutes & ":" & Seconds '& ":" & MilliSec
End Sub

Public Function CheckItem()
    If chkDisBuff.value = 1 Then
        DisBuffer = False
        ProgScan.TextStyle = CustomText
    Else
        DisBuffer = True
        ProgScan.TextStyle = PBPercentage
    End If
End Function

Private Sub SelectAll()
    Dim I As Integer
    
    With lstDetection.ListItems
        For I = 1 To .count
            .Item(I).Checked = True
        Next I
    End With
End Sub

Private Sub Unselect()
    Dim I As Integer
    
    With lstDetection.ListItems
        For I = 1 To .count
            .Item(I).Checked = False
        Next I
    End With
End Sub

Private Function CheckVirus()
    Dim I As Double
    
    With lstDetection.ListItems
        For I = 1 To .count
            If .Item(I).Checked = True Then
                CheckVirus = CheckVirus + 1
            End If
        Next I
    End With
End Function

Public Sub CleanVirus()
    On Error Resume Next
    
    Dim strClean As String
    Dim I As Long, lRet As Long
    
    With lstDetection.ListItems
        For I = 1 To .count
            If .Item(I).Checked = True Then
                strClean = .Item(I).SubItems(1)
                SetFileAttributes strClean, FILE_ATTRIBUTE_NORMAL
                Tunggu 1
                DoEvents
                LogFile "Cure    " & strClean
                DeleteIt (strClean)
                If lRet <> 0 Then
                    .Item(I).Checked = False
                End If
                .Item(I).Checked = False
                CleanVirus
                VirusCleaned = VirusCleaned + 1
                lblVirusClean.Caption = ": " & VirusCleaned
                txtStatus.ForeColor = &H80000008
                txtStatus.Text = "STATUS : Object Cleaned."
                .Remove (I)
                Exit Sub
            End If
        Next I
    End With
End Sub

Private Sub Quarantine()
On Error Resume Next
    
    Dim nama, Exten As String
    Dim I As Long
    Dim strFile As String, strName As String
    
    With lstDetection.ListItems
        For I = 1 To .count
            strFile = .Item(I).SubItems(1)
            txtStatus.ForeColor = &H80000008
            txtStatus.Text = "STATUS : Quarantine object"
            If .Item(I).Checked Then
                nama = GetFileName(strFile)
                Exten = Right$(strFile, 3)
                SetFileAttributes nama, FILE_ATTRIBUTE_NORMAL
                Tunggu 1
                DoEvents
                TerminateExeName strFile
                DocFix (WhereMine)
                LogFile "Quarantine    " & strFile
                If Seal.EncodeFile(strFile, App.path & "\Quarantine\" & nama & "." & Exten & ".vir") = False Then
                    MsgBox "Cleaning Virus Failed  !", vbOKOnly, APP_PROGRAM
                End If
                Open (strFile) For Output As #1
                Close (1)
                Kill (strFile)
                VirusCleaned = VirusCleaned + 1
                lblVirusClean.Caption = ": " & VirusCleaned
                txtStatus.Text = "STATUS : Object (s) has been added to quarantine folder."
                txtStatus.ForeColor = &H80000008
                .Remove I
                Exit Sub
            End If
        Next I
    End With
End Sub

Private Sub ClearLabel()
    picIconP32.Cls
    lblDescription.Caption = ""
    lblCompany.Caption = ""
    lblPath.Caption = ""
    lblFile.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
End Sub

Private Sub usb_Click()
MsgBox "Feature Not build Yet", vbOKOnly, APP_PROGRAM
End Sub

Private Sub updd_Click()
' here will be the link for the update file on my server
' hope it works dude...

ShellExecute 0, "open", "http://321321.atwebpages.com/Download/TCM.rar", vbNullString, vbNullString, 1    ' here is the style
End Sub

Private Sub usbb_Click()
ShellExecute Me.hWnd, vbNullString, nPath(App.path) & "\Usb Disinfector .exe", vbNullString, "C:\", 1
End Sub
