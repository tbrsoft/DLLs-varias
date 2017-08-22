VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monoton Player"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   518
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTabEQ 
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   300
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   259
      TabIndex        =   13
      Top             =   5970
      Visible         =   0   'False
      Width           =   3885
      Begin VB.CheckBox chkEQ 
         Caption         =   "Enabled"
         Height          =   195
         Left            =   720
         TabIndex        =   35
         Top             =   0
         Width           =   1005
      End
      Begin prjMonoPlayer.ctlEBSlider sldPreamp 
         Height          =   915
         Left            =   360
         TabIndex        =   14
         Top             =   510
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   1614
         Min             =   -120
         Max             =   120
         Value           =   0
         Orientation     =   1
      End
      Begin prjMonoPlayer.ctlEBSlider sldEQBand 
         Height          =   915
         Index           =   0
         Left            =   1440
         TabIndex        =   16
         Top             =   510
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   1614
         Min             =   -120
         Max             =   120
         Value           =   0
         Orientation     =   1
      End
      Begin prjMonoPlayer.ctlEBSlider sldEQBand 
         Height          =   915
         Index           =   1
         Left            =   1950
         TabIndex        =   17
         Top             =   510
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   1614
         Min             =   -120
         Max             =   120
         Value           =   0
         Orientation     =   1
      End
      Begin prjMonoPlayer.ctlEBSlider sldEQBand 
         Height          =   915
         Index           =   2
         Left            =   2475
         TabIndex        =   18
         Top             =   510
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   1614
         Min             =   -120
         Max             =   120
         Value           =   0
         Orientation     =   1
      End
      Begin prjMonoPlayer.ctlEBSlider sldEQBand 
         Height          =   915
         Index           =   3
         Left            =   2985
         TabIndex        =   19
         Top             =   510
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   1614
         Min             =   -120
         Max             =   120
         Value           =   0
         Orientation     =   1
      End
      Begin prjMonoPlayer.ctlEBSlider sldEQBand 
         Height          =   915
         Index           =   4
         Left            =   3510
         TabIndex        =   20
         Top             =   510
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   1614
         Min             =   -120
         Max             =   120
         Value           =   0
         Orientation     =   1
      End
      Begin VB.Label lblDB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   5
         Left            =   3420
         TabIndex        =   34
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblDB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   2880
         TabIndex        =   33
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblDB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   2400
         TabIndex        =   32
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblDB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   1890
         TabIndex        =   31
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblDB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   1350
         TabIndex        =   30
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblDB 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   270
         TabIndex        =   29
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblDBID 
         AutoSize        =   -1  'True
         Caption         =   "0 dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   810
         TabIndex        =   28
         Top             =   870
         Width           =   285
      End
      Begin VB.Label lblDBID 
         AutoSize        =   -1  'True
         Caption         =   "-12 dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   780
         TabIndex        =   27
         Top             =   1260
         Width           =   405
      End
      Begin VB.Label lblDBID 
         AutoSize        =   -1  'True
         Caption         =   "+12 dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   720
         TabIndex        =   26
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lblEQBand 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 Hz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   3420
         TabIndex        =   25
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label lblEQBand 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 Hz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   2910
         TabIndex        =   24
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label lblEQBand 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 Hz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   2430
         TabIndex        =   23
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label lblEQBand 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 Hz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   1920
         TabIndex        =   22
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label lblEQBand 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 Hz"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   1395
         TabIndex        =   21
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label lblPreamp 
         AutoSize        =   -1  'True
         Caption         =   "Preamp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   15
         Top             =   1440
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdNext 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2460
      Picture         =   "frmPlayer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4980
      Width           =   465
   End
   Begin VB.CommandButton cmdPrev 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "frmPlayer.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4980
      Width           =   465
   End
   Begin VB.CommandButton cmdPLRem 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      TabIndex        =   47
      Top             =   3690
      Width           =   465
   End
   Begin VB.CommandButton cmdPLAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4950
      TabIndex        =   46
      Top             =   3690
      Width           =   465
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   3525
      Left            =   4770
      TabIndex        =   45
      Top             =   90
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6218
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   3439
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Length"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.PictureBox picTabFX 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   300
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   36
      Top             =   5880
      Visible         =   0   'False
      Width           =   3795
      Begin VB.CheckBox chkFXEcho 
         Caption         =   "Echo"
         Height          =   195
         Left            =   270
         TabIndex        =   50
         Top             =   1440
         Width           =   2445
      End
      Begin VB.CheckBox chkFXHighPass 
         Caption         =   "High Pass Filter"
         Height          =   195
         Left            =   270
         TabIndex        =   43
         Top             =   630
         Width           =   1455
      End
      Begin MSComctlLib.Slider sldFXLowPass 
         Height          =   270
         Left            =   900
         TabIndex        =   39
         Top             =   360
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   476
         _Version        =   393216
         Min             =   30
         Max             =   16000
         SelStart        =   10000
         TickStyle       =   3
         Value           =   10000
      End
      Begin VB.CheckBox chkFXFlanger 
         Caption         =   "Flanger"
         Height          =   195
         Left            =   270
         TabIndex        =   38
         Top             =   1170
         Width           =   1635
      End
      Begin VB.CheckBox chkFXLowPass 
         Caption         =   "Low Pass Filter"
         Height          =   195
         Left            =   270
         TabIndex        =   37
         Top             =   90
         Width           =   1455
      End
      Begin MSComctlLib.Slider sldFXHighPass 
         Height          =   270
         Left            =   900
         TabIndex        =   42
         Top             =   900
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   476
         _Version        =   393216
         Min             =   30
         Max             =   16000
         SelStart        =   1000
         TickStyle       =   3
         Value           =   1000
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cutoff:"
         Height          =   195
         Left            =   360
         TabIndex        =   44
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cutoff:"
         Height          =   195
         Left            =   360
         TabIndex        =   40
         Top             =   360
         Width           =   525
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7500
      Top             =   5010
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTabTags 
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   135
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   11
      Top             =   5970
      Width           =   4245
      Begin MSComctlLib.ListView lvwTags 
         Height          =   1545
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   2725
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   1985
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   4498
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tabCtrl 
      Height          =   2175
      Left            =   60
      TabIndex        =   10
      Top             =   5520
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   3836
      ShowTips        =   0   'False
      TabMinWidth     =   2059
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tags"
            Key             =   "TAGS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Equalizer"
            Key             =   "EQ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Effects"
            Key             =   "FX"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   375
      Left            =   3180
      Picture         =   "frmPlayer.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4980
      Width           =   555
   End
   Begin prjMonoPlayer.ucSlider sldVolume 
      Height          =   240
      Left            =   5325
      Top             =   5280
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   423
      SliderIcon      =   "frmPlayer.frx":03DE
      Orientation     =   0
      RailPicture     =   "frmPlayer.frx":0538
      RailStyle       =   2
      Max             =   100
      Value           =   100
   End
   Begin VB.PictureBox picTimeVis 
      BackColor       =   &H00000000&
      Height          =   3645
      Left            =   165
      ScaleHeight     =   3585
      ScaleWidth      =   4185
      TabIndex        =   6
      Top             =   90
      Width           =   4245
      Begin VB.PictureBox picVis 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3285
         Left            =   15
         ScaleHeight     =   219
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   275
         TabIndex        =   7
         Top             =   270
         Width           =   4125
      End
      Begin VB.Label lblVisName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bars"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   41
         Top             =   45
         Width           =   315
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   960
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
   End
   Begin prjMonoPlayer.MsgScroller scrlTitle 
      Height          =   285
      Left            =   5235
      Top             =   4650
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   503
      BackColor       =   0
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Seperator       =   "  ·  "
      AutoScroll      =   -1  'True
   End
   Begin prjMonoPlayer.ucSlider sldPosition 
      Height          =   240
      Left            =   15
      Top             =   4680
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   423
      SliderIcon      =   "frmPlayer.frx":0554
      Orientation     =   0
      RailPicture     =   "frmPlayer.frx":0666
      RailStyle       =   3
   End
   Begin VB.CommandButton cmdStop 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1815
      Picture         =   "frmPlayer.frx":0682
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4980
      Width           =   465
   End
   Begin VB.CommandButton cmdPause 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1275
      Picture         =   "frmPlayer.frx":07CC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4980
      Width           =   465
   End
   Begin VB.CommandButton cmdPlay 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   735
      Picture         =   "frmPlayer.frx":0916
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4980
      Width           =   465
   End
   Begin prjMonoPlayer.ucSlider sldPan 
      Height          =   240
      Left            =   6945
      Top             =   5280
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   423
      SliderIcon      =   "frmPlayer.frx":0A60
      Orientation     =   0
      RailPicture     =   "frmPlayer.frx":0BBA
      RailStyle       =   2
      Min             =   -2000
      Max             =   2000
   End
   Begin VB.Label lblBitrate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "bps"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   5790
      TabIndex        =   5
      Top             =   5010
      Width           =   210
   End
   Begin VB.Label lblChannels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   7320
      TabIndex        =   4
      Top             =   5010
      Width           =   45
   End
   Begin VB.Label lblSamplerate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "kHz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   6495
      TabIndex        =   3
      Top             =   5010
      Width           =   240
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IDSOutCallback

Private Enum VISUALISATION
    VIS_BARS
    VIS_OSC
    VIS_PEAKS
    VIS_NONE
    VIS_COUNT
End Enum

Private Const INTERVAL_TIME         As Long = 200
Private Const INTERVAL_VIS          As Long = 50

Private clsLoader       As tbrDX8.MonotonLoader
Private clsPlayer       As tbrDX8.SoundOut
Private clsEqualizer    As tbrDX8.Equalizer
Private clsLowPass      As tbrDX8.IIRFilter
Private clsHighPass     As tbrDX8.IIRFilter

Private WithEvents clsTimerPos      As clsTimer
Attribute clsTimerPos.VB_VarHelpID = -1
Private WithEvents clsTimerVis      As clsTimer
Attribute clsTimerVis.VB_VarHelpID = -1

Private udeVis              As VISUALISATION
Private strCurrentFile      As String
Private blnDontMovePos      As Boolean
Private lngPlaylistIndex    As Long

' *********************************************
' * Visualisations
' *********************************************

Private Sub clsTimerVis_Timer()
    Dim intSamples(FFT_SAMPLES - 1) As Integer

    If clsPlayer.bitspersample = 16 Then
        clsPlayer.CaptureSamples VarPtr(intSamples(0)), FFT_SAMPLES * 2

        Select Case udeVis
            Case VIS_BARS
                modDraw.DrawFrequencies intSamples, picVis
            Case VIS_OSC
                modDraw.DrawOsc intSamples, picVis
            Case VIS_PEAKS
                modDraw.DrawPeaks intSamples, picVis
            Case VIS_NONE
                picVis.Cls
        End Select
    End If
End Sub

' ************************************************
' * Player events
' ************************************************

Private Sub IDSOutCallback_OnEndOfStream( _
    sndout As tbrDX8.SoundOut _
)

    Debug.Print "End Of Stream"

    PlayListPlayNext
End Sub

Private Sub IDSOutCallback_OnStatusChanged( _
    sndout As tbrDX8.SoundOut, _
    ByVal udeStat As tbrDX8.DS_PlayState _
)

    Debug.Print "Status: ";

    Select Case udeStat

        Case STAT_PAUSING
            Debug.Print "Pause"
            clsTimerPos.Enabled = False
            clsTimerVis.Enabled = False

        Case STAT_PLAYING
            Debug.Print "Play"
            clsTimerPos.Enabled = True
            clsTimerVis.Enabled = True

        Case STAT_STOPPED
            Debug.Print "Stop"
            clsTimerPos.Enabled = False
            clsTimerVis.Enabled = False

    End Select
End Sub

Private Sub IDSOutCallback_Samples( _
    sndout As tbrDX8.SoundOut, _
    intSamples() As Integer, _
    ByVal datalength As Long, _
    ByVal channels As Integer _
)

    datalength = datalength / 2 - 1

    clsDSP.ChangeVolume intSamples, datalength, sldPreamp.Value / 10, VOL_DECIBEL

    If chkFXEcho.Value Then
        modDSP.DSPEcho intSamples, datalength
    End If

    If chkFXLowPass.Value Then
        clsLowPass.ProcessSamples intSamples, datalength
    End If

    If chkFXHighPass.Value Then
        clsHighPass.ProcessSamples intSamples, datalength
    End If

    If chkFXFlanger.Value Then
        modDSP.DSPFlange intSamples, datalength
    End If

    If chkEQ.Value Then
        clsEqualizer.ProcessSamples intSamples, datalength
    End If
End Sub

' ************************************************
' * Filehandling
' ************************************************

Private Sub ShowTags()
    Dim clsTag  As tbrDX8.StreamTag

    With lvwTags.ListItems
        .Clear

        For Each clsTag In clsStream.Info.Tags
            .Add(Text:=clsTag.TagName).SubItems(1) = clsTag.TagValue
        Next
    End With

    lblBitrate.Caption = Fix(clsStream.Info.Bitrate / 1000) & " bps"
    lblSamplerate.Caption = Fix(clsStream.Info.samplerate / 1000) & " kHz"
    lblChannels.Caption = Choose(clsStream.Info.channels, "mono", "stereo")
End Sub

' *********************************************
' * Playlist
' *********************************************

Private Function PlaylistOpenFile( _
    ByVal Index As Long _
) As Boolean

    On Error GoTo ErrorHandler

    Dim i       As Long
    Dim strFile As String

    strFile = PlaylistFileGetPath(Index)

    If Not clsStream Is Nothing Then
        clsStream.CloseSource
        clsPlayer.StopPlay
    End If

    Set clsStream = Nothing

    Set clsStream = StreamFromExt(strFile)
    If clsStream Is Nothing Then
        Debug.Print "Format nicht unterstützt"
        Exit Function
    End If

    If clsStream.OpenSource(strFile) <> STREAM_OK Then
        Debug.Print "Datei konnte nicht geöffnet werden"
        Exit Function
    End If

    sldPosition.Max = clsStream.Info.Duration
    PlaylistFileSetTime Index, clsStream.Info.Duration

    ShowTags

    If Not clsPlayer.Initialize(clsStream, Me) Then
        Debug.Print "DSOut konnte nicht initialisiert werden"
        Exit Function
    End If

    ' Player settings
    clsPlayer.Volume = sldVolume.Value
    clsPlayer.Pan = sldPan.Value

    ' filename
    strCurrentFile = strFile

    scrlTitle.Clear
    scrlTitle.AddItem GetFilename(strFile) & " (" & FmtTime(clsStream.Info.Duration / 1000) & ")"

    ' DSP settings
    clsEqualizer.samplerate = clsPlayer.samplerate
    For i = 0 To 4
        sldEQBand_Changed CInt(i)
    Next

    Set clsLowPass = clsDSP.CreateBiquadIIR(IIR_LOW_PASS, 0, sldFXLowPass.Value, clsPlayer.samplerate, 1.5)
    Set clsHighPass = clsDSP.CreateBiquadIIR(IIR_HIGH_PASS, 0, sldFXHighPass.Value, clsPlayer.samplerate, 1.5)

    modDSP.DSPEchoSettings clsPlayer.samplerate, 500, 0.4

    If Not clsPlayer.Play Then
        Debug.Print "Could not play!"
        Exit Function
    End If

    lngPlaylistIndex = Index

    PlaylistOpenFile = True
    Exit Function

ErrorHandler:
    clsStream.CloseSource
    clsPlayer.StopPlay
End Function

Public Sub PlayListPlayPrev()
    lngPlaylistIndex = lngPlaylistIndex - 1
    If lngPlaylistIndex < 1 Then
        lngPlaylistIndex = 1
        Exit Sub
    End If

    If Not PlaylistOpenFile(lngPlaylistIndex) Then
        PlayListPlayNext
    End If
End Sub

Public Sub PlayListPlayNext()
    lngPlaylistIndex = lngPlaylistIndex + 1
    If lngPlaylistIndex > lvwFiles.ListItems.Count Then
        lngPlaylistIndex = 1
    End If

    If Not PlaylistOpenFile(lngPlaylistIndex) Then
        PlayListPlayNext
        DoEvents
    End If
End Sub

Private Sub PlaylistClear()
    lvwFiles.ListItems.Clear
End Sub

Private Sub PlaylistAddFile( _
    ByVal file As String _
)

    If InStr(1, _
             Join(GetAllExtensions, ";"), _
             GetExtension(file), _
             vbTextCompare) < 1 Then

        Exit Sub
    End If

    With lvwFiles.ListItems.Add(Text:=GetFilename(file))
        .Tag = file
    End With
End Sub

Private Sub PlaylistFileSetTime( _
    ByVal Index As Long, _
    ByVal ms As Long _
)

    With lvwFiles.ListItems(Index)
        .SubItems(1) = FmtTime(ms / 1000)
    End With
End Sub

Private Sub PlaylistFileSetTitle( _
    ByVal Index As Long, _
    ByVal title As String _
)

    With lvwFiles.ListItems(Index)
        .Text = title
    End With
End Sub

Private Function PlaylistFileGetPath( _
    ByVal Index As Long _
) As String

    PlaylistFileGetPath = lvwFiles.ListItems(Index).Tag
End Function


' *********************************************
' * UI
' *********************************************

Private Sub clsTimerPos_Timer()
    lblTime.Caption = FmtTime(clsPlayer.position / 1000)

    If Not blnDontMovePos Then
        If Not clsPlayer.position > sldPosition.Max Then
            sldPosition.Value = clsPlayer.position
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
    Dim strFiles()  As String
    Dim strPath     As String
    Dim i           As Long

    dlg.FileName = vbNullString
    dlg.Filter = Join(GetAllExtensions, "; ") & "|*." & Join(GetAllExtensions, ";*.")
    dlg.Flags = cdlOFNAllowMultiselect Or _
                cdlOFNExplorer Or _
                cdlOFNLongNames

    dlg.MaxFileSize = &H7FFF

    dlg.ShowOpen
    If dlg.FileName = vbNullString Then Exit Sub

    PlaylistClear

    If InStr(dlg.FileName, Chr$(0)) > 0 Then
        strFiles = Split(dlg.FileName, Chr$(0))
        strPath = AddSlash(strFiles(0))

        For i = 1 To UBound(strFiles)
            strFiles(i - 1) = strPath & strFiles(i)
        Next

        ReDim Preserve strFiles(UBound(strFiles) - 1) As String
    Else
        ReDim strFiles(0) As String
        strFiles(0) = dlg.FileName
    End If

    For i = 0 To UBound(strFiles)
        PlaylistAddFile strFiles(i)
    Next

    lngPlaylistIndex = 0
    PlayListPlayNext
End Sub

Private Sub cmdPause_Click()
    clsPlayer.Pause
End Sub

Private Sub cmdPlay_Click()
    If Not clsPlayer.Play Then
        MsgBox "Could not play!", vbExclamation
    End If
End Sub

Private Sub cmdNext_Click()
    PlayListPlayNext
End Sub

Private Sub cmdPrev_Click()
    PlayListPlayPrev
End Sub

Private Sub cmdStop_Click()
    clsPlayer.StopPlay
    clsPlayer.SeekTo 0, SEEK_PERCENT
End Sub

Private Sub Form_Load()
    ProcessThreadPrioritySet 2

    AddStream New tbrDX8.StreamAPE
    AddStream New tbrDX8.StreamCDA
    AddStream New tbrDX8.StreamMP3
    AddStream New tbrDX8.StreamOGG
    AddStream New tbrDX8.StreamWAV
    AddStream New tbrDX8.StreamWMA

    Set clsEqualizer = New tbrDX8.Equalizer
    Set clsLoader = New tbrDX8.MonotonLoader
    Set clsDSP = New tbrDX8.SignalProcessor

    Set clsTimerPos = New clsTimer
    Set clsTimerVis = New clsTimer

    clsTimerPos.Interval = INTERVAL_TIME
    clsTimerVis.Interval = INTERVAL_VIS

    clsEqualizer.samplerate = 44100
    clsEqualizer.SetBandCount 5

    If Not clsLoader.Initialize(44100, 2, 16) Then
        MsgBox "Konnte DirectSound 8 (Default Device) nicht initialisieren!", vbExclamation
        Unload Me
    End If

    Set clsPlayer = clsLoader.CreateSoundOut()
    If clsPlayer Is Nothing Then
        MsgBox "Could not create SoundOut!", vbExclamation
        Unload Me
    End If

    clsPlayer.Force16Bit = True

    clsPlayer.VolumeUnit = UNIT_LINEAR
End Sub

Private Sub Form_Unload( _
    Cancel As Integer _
)

    Dim i   As Long

    Set clsPlayer = Nothing
    Set clsLoader = Nothing

    For i = 0 To lngStreamCnt - 1
        Set clsStreams(i) = Nothing
    Next
End Sub

Private Sub lvwFiles_DblClick()
    If Not PlaylistOpenFile(lvwFiles.SelectedItem.Index) Then
        MsgBox "Could not open the file!", vbExclamation
    End If
End Sub

Private Sub lvwFiles_OLEDragDrop( _
    Data As MSComctlLib.DataObject, _
    Effect As Long, _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    Y As Single _
)

    Dim i   As Long

    With Data.Files
        For i = 1 To .Count
            If FileExists(.Item(i)) Then
                PlaylistAddFile .Item(i)
            End If
        Next
    End With
End Sub

Private Sub picVis_Click()
    udeVis = udeVis + 1
    If udeVis = VIS_COUNT Then udeVis = 0

    Select Case udeVis
        Case VIS_BARS
            picVis.AutoRedraw = True
            lblVisName.Caption = "Bars"
        Case VIS_PEAKS
            picVis.AutoRedraw = True
            lblVisName.Caption = "Peaks"
        Case VIS_NONE
            picVis.AutoRedraw = True
            lblVisName.Caption = "None"
        Case VIS_OSC
            picVis.AutoRedraw = False
            lblVisName.Caption = "Osc"
    End Select
End Sub

Private Sub sldEQBand_Changed( _
    Index As Integer _
)

    Dim lngFreq     As Long
    Dim dblBW       As Double
    Dim dblBandFreq As Double

    If clsPlayer.samplerate > 0 Then

        ' max. frequency: samplerate / 3
        lngFreq = clsPlayer.samplerate / 3

        ' bandwidth for current band
        dblBW = Log(lngFreq / 80#) / 4

        With clsEqualizer
            ' center of the band
            dblBandFreq = 80# * (lngFreq / 80#) ^ (Index / 4#)
            .SetBandValues Index, sldEQBand(Index).Value / 10, dblBandFreq, dblBW
        End With

        If dblBandFreq > 1000 Then
            lblEQBand(Index).Caption = Fix(dblBandFreq / 1000) & " kHz"
        Else
            lblEQBand(Index).Caption = Fix(dblBandFreq) & " Hz"
        End If
    End If

    lblDB(Index + 1).Caption = Fix(sldEQBand(Index).Value / 10) & " dB"
End Sub

Private Sub sldFXHighPass_Click()
    Set clsHighPass = clsDSP.CreateBiquadIIR(IIR_HIGH_PASS, 0, sldFXHighPass.Value, clsPlayer.samplerate, 1.5)
End Sub

Private Sub sldFXLowPass_Change()
    Set clsLowPass = clsDSP.CreateBiquadIIR(IIR_LOW_PASS, 0, sldFXLowPass.Value, clsPlayer.samplerate, 1.5)
End Sub

Private Sub sldPan_Change()
    clsPlayer.Pan = sldPan.Value
End Sub

Private Sub sldPosition_MouseDown( _
    Shift As Integer _
)

    blnDontMovePos = True
End Sub

Private Sub sldPosition_MouseUp( _
    Shift As Integer _
)

    blnDontMovePos = False

    If clsStream Is Nothing Then Exit Sub

    If Not clsPlayer.SeekTo(sldPosition.Value / 1000, SEEK_SECONDS) Then
        Debug.Print "Konnte nicht seeken"
    End If
End Sub

Private Sub sldPreamp_Changed()
    lblDB(0).Caption = Fix(sldPreamp.Value / 10) & " dB"
End Sub

Private Sub sldVolume_Change()
    clsPlayer.Volume = sldVolume.Value
End Sub

Private Sub tabCtrl_Click()
    Select Case tabCtrl.SelectedItem.Key
        Case "TAGS"
            picTabEQ.Visible = False
            picTabTags.Visible = True
            picTabFX.Visible = False
        Case "EQ"
            picTabEQ.Visible = True
            picTabTags.Visible = False
            picTabFX.Visible = False
        Case "FX"
            picTabEQ.Visible = False
            picTabTags.Visible = False
            picTabFX.Visible = True
    End Select
End Sub

Private Sub cmdPLRem_Click()
    Dim i   As Long

    For i = lvwFiles.ListItems.Count To 1 Step -1
        If lvwFiles.ListItems(i).Selected Then
            lvwFiles.ListItems.Remove i
            If i <= lngPlaylistIndex Then
                If lngPlaylistIndex > 1 Then
                    lngPlaylistIndex = lngPlaylistIndex - 1
                End If
            End If
        End If
    Next
End Sub

Private Sub cmdPLAdd_Click()
    Dim strFiles()  As String
    Dim strPath     As String
    Dim i           As Long

    dlg.FileName = vbNullString
    dlg.Filter = Join(GetAllExtensions, "; ") & "|*." & Join(GetAllExtensions, ";*.")
    dlg.Flags = cdlOFNAllowMultiselect Or _
                cdlOFNExplorer Or _
                cdlOFNLongNames

    dlg.MaxFileSize = &H7FFF

    dlg.ShowOpen
    If dlg.FileName = vbNullString Then Exit Sub

    If InStr(dlg.FileName, Chr$(0)) > 0 Then
        strFiles = Split(dlg.FileName, Chr$(0))
        strPath = AddSlash(strFiles(0))

        For i = 1 To UBound(strFiles)
            strFiles(i - 1) = strPath & strFiles(i)
        Next

        ReDim Preserve strFiles(UBound(strFiles) - 1) As String
    Else
        ReDim strFiles(0) As String
        strFiles(0) = dlg.FileName
    End If

    For i = 0 To UBound(strFiles)
        PlaylistAddFile strFiles(i)
    Next
End Sub
