VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Take a photo.........."
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Left            =   4800
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   1440
   End
   Begin VB.CommandButton com 
      BackColor       =   &H0000FFFF&
      Caption         =   "Take photo"
      Height          =   735
      Left            =   1680
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Image cow 
      Height          =   495
      Left            =   360
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2835
      Left            =   1200
      Picture         =   "Form1.frx":030A
      Top             =   120
      Width           =   3345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
   '   PlaySound App.Path & "\camera.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC

Private Sub com_Click()
MsgBox " go a bit to the right to fit in the picture , smile ....", , "Take a photo"
  PlaySound App.Path & "\camera.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
Prog.Visible = True
com.Visible = False
Image1.Visible = False
Timer2.Interval = 10
End Sub

Private Sub Form_Load()
Prog.Visible = False
cow.Top = -(cow.Height)
cow.Left = (Me.ScaleWidth) / 2 - cow.Width
End Sub

Private Sub Timer1_Timer()
cow.Picture = LoadPicture(App.Path & "\cow.jpg")
cow.Top = cow.Top + 100
If cow.Top > (Me.ScaleHeight) / 2 - cow.Height / 2 Then Call soundd: Timer1.Interval = 0
End Sub

Private Sub Timer2_Timer()
If Prog.Value = 99 Then
Prog.Visible = False
Timer1.Interval = 100
Call continue
Exit Sub
End If
Prog.Value = Prog.Value + 1
End Sub


Sub continue()
Timer2.Interval = 0
'Timer1.Interval = 1000

End Sub

Sub soundd()
  PlaySound App.Path & "\cow.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
End Sub
