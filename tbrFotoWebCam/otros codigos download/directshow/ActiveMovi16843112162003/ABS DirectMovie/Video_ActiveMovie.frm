VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Video_ActiveMovie 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VideoWindow"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   574
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   486
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Info"
      Height          =   855
      Left            =   360
      TabIndex        =   13
      Top             =   7200
      Width           =   4455
      Begin VB.Timer RefreshTimer 
         Interval        =   250
         Left            =   4080
         Top             =   120
      End
      Begin VB.Label CurrentPos_l 
         Caption         =   "Current Pos:"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Length_l 
         Caption         =   "Length:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Controls"
      Height          =   1455
      Left            =   360
      TabIndex        =   8
      Top             =   4440
      Width           =   4455
      Begin VB.CheckBox Ratio_c 
         Caption         =   "Maintain Aspect Ratio"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox FullScreen_c 
         Caption         =   "Run Full Screen"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Path_t 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "C:\Tvm\Station\Media\jump.avi"
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton Play_but 
         Caption         =   "Play"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Stop_but 
         Caption         =   "Stop"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Pause_but 
         Caption         =   "Pause"
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Audio"
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   6000
      Width           =   4455
      Begin MSComctlLib.Slider Volume_s 
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         Min             =   -4000
         Max             =   0
         TickFrequency   =   250
      End
      Begin MSComctlLib.Slider Balance_s 
         Height          =   495
         Left            =   2400
         TabIndex        =   6
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         Min             =   -5000
         Max             =   5000
         TickFrequency   =   500
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Balance"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Volume"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Timer StateTimer 
      Interval        =   250
      Left            =   5760
      Top             =   1080
   End
   Begin VB.PictureBox Video 
      BackColor       =   &H00000000&
      Height          =   3615
      Left            =   360
      ScaleHeight     =   3555
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "    Coded By Fade"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      ToolTipText     =   "Coded By Fade(Amit Ben-Shahar)"
      Top             =   8160
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "A timer that captures object state"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   $"Video_ActiveMovie.frx":0000
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Video_ActiveMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Paused
Const NormalWidth = 5280


Private Sub Form_Load()
    Me.Width = NormalWidth
End Sub

' **
' **  Control Buttons
' **
    
Private Sub Play_But_Click() ' Play
    If Paused Then ' Check if paused
        ActiveMovieControl.PlayActiveMovie
    Else    ' if not, new content
        DontMaintainRatio = (Ratio_c.Value = 0)
        RunFullScreen = (FullScreen_c.Value = 1)
        ActiveMovieControl.RunVideoContent Path_t.Text, DontMaintainRatio, RunFullScreen
    End If
End Sub

Private Sub Stop_But_Click() ' Stop
        ' Setting flag
    Paused = False
    ' ****************
    ActiveMovieControl.StopActiveMovie
End Sub

Private Sub Pause_But_Click()
    ' Setting Flag
    Paused = True
    ' ---------------
    ActiveMovieControl.PauseActiveMovie
End Sub

' **
' ** Audio Control Slides
' **
' ** Note: The 'Click' event in slides will only capture 'drags'
' **    that finishes inside the control's area, to get events
' **    during the drag use the 'mouseMove' event for smooth handling

Private Sub Volume_s_Click()
    ActiveMovieControl.SetActiveMovieVolume Volume_s.Value
End Sub


Private Sub Balance_s_Click()
    ActiveMovieControl.SetActiveMovieBalance Balance_s.Value
End Sub


' **
' ** Timer Events
' **

Private Sub RefreshTimer_Timer()
    If ActiveMovieControl.VideoRunning Then
        Length_l.Caption = "Length: " & ActiveMovieControl.GetVideoLength
        CurrentPos_l.Caption = "Current Pos: " & ActiveMovieControl.GetVideoPos
    End If
End Sub

Private Sub StateTimer_Timer()
    ActiveMovieControl.ActiveMovieTimerEvent
End Sub

' **
' **  Video Finished Event
' **

Public Sub VideoFinishedEvent()
    CurrentPos_l.Caption = "Video Finished!"
End Sub

