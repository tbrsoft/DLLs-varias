VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "game"
   ClientHeight    =   4704
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6516
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4704
   ScaleWidth      =   6516
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   2
      Top             =   0
      Width           =   612
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1092
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   4692
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   612
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1332
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Sky = &HFFFFC0
Const Grass = &H7FFF00 ' BGR
Const Expl = &H7F00FF
Const SPD = 0.25
Const GWIDTH = 2

Dim IsOK As Boolean, Speed As Long
Dim cWidth As Long, cHeight As Long
Dim zWidth As Long, zHeight As Long, q() As Long

Private Sub Command1_Click()
  Init
End Sub

Private Sub Form_Activate()
  MsgBox "Click on Command1, then click on the green", vbInformation, "game"
End Sub

Private Sub Form_Load()
  cWidth = (Screen.Width - (Screen.Width / 5)) / Screen.TwipsPerPixelX
  cHeight = (Screen.Height - (Screen.Height / 5)) / Screen.TwipsPerPixelY
  Picture1.BackColor = Sky
  Form1.Width = cWidth * Screen.TwipsPerPixelX
  Form1.Height = cHeight * Screen.TwipsPerPixelY
  Frame1.Width = Form1.ScaleWidth
  Frame1.Top = Form1.ScaleHeight - Frame1.Height
  Picture1.Width = Form1.ScaleWidth
  Picture1.Height = Frame1.Top - 1
  Form1.Left = (Screen.Width - Form1.Width) / 2.25
  Form1.Top = (Screen.Height - Form1.Height) / 3
  Picture1.ScaleMode = 3
  zWidth = Picture1.ScaleWidth
  zHeight = Picture1.ScaleHeight
  ReDim q(zWidth)
End Sub

Private Sub Init()
  Dim t As Long, z As Long, i As Long
  Randomize Timer
  SetSpeed
  Picture1.Cls
  SetColor Picture1, Grass, 1
  z = zHeight / 1.5
  For t = 0 To zWidth
    i = CLng((Rnd * 6) - 3)
    If z >= zHeight Then
      z = z - Abs(i)
    ElseIf z <= zHeight / 3 Then
      z = z + Abs(i)
    Else
      z = z + i
    End If
    q(t) = z
    DrawLine Picture1, t, q(t), t, zHeight
  Next t
  IsOK = True
End Sub

Private Sub CircleExp(ByVal X As Long, ByVal Y As Long)
If IsOK Then
  IsOK = False
  Dim t As Long, tt As Long, StartPos As Long, EndPos As Long, z As Long, Max As Byte
  SetColor Picture1, Expl, GWIDTH
  For t = 1 To zWidth / 20
    DrawCircle Picture1, X, Y, t
    If t Mod Speed = 0 Then Picture1.Refresh
  Next t
  Picture1.Refresh
  SetColor Picture1, Sky
  For t = 1 To zWidth / 20
    DrawCircle Picture1, X, Y, t
    If t Mod Speed = 0 Then Picture1.Refresh
  Next t
  Picture1.Refresh
  StartPos = (X - (zWidth / 20)) - (GWIDTH / 2)
  EndPos = (X + (zWidth / 20)) + (GWIDTH / 2)
  If StartPos < 0 Then StartPos = 0
  If EndPos > zWidth Then EndPos = zWidth
  For t = StartPos To EndPos
    z = 0
    For tt = zHeight / 3 To zHeight
       If GetPixel(Picture1, t, tt) = Grass Then z = z + 1
    Next tt
    q(t) = zHeight - z
    SetColor Picture1, Sky, 1
    DrawLine Picture1, t, 0, t, q(t) - 1
    SetColor Picture1, Grass
    DrawLine Picture1, t, q(t), t, zHeight
  Next t
  Picture1.Refresh
  IsOK = True
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  CircleExp X, Y
End Sub

Private Sub Slower(qq As Single)
  Dim t As Long, qqq As Single
  If Speed > 1 Then Speed = Speed - 1
  Do
    qq = Timer
    qqq = qq
    SetColor Picture1, Expl
    For t = 1 To zWidth / 20
      DrawCircle Picture1, (zWidth / 20), (zWidth / 20), t
      If t Mod Speed = 0 Then Picture1.Refresh
    Next t
    Picture1.Refresh
    SetColor Picture1, Sky
    For t = 1 To zWidth / 20
      DrawCircle Picture1, (zWidth / 20), (zWidth / 20), t
      If t Mod Speed = 0 Then Picture1.Refresh
    Next t
    Picture1.Refresh
    qq = Timer - qq
    If (qq < SPD) Then Speed = Speed - 1
  Loop Until (qq >= SPD) Or (Speed < 1)
  If ((SPD - qqq) < (qq - SPD)) Or (Speed < 1) Then Speed = Speed + 1
End Sub

Private Sub Faster(qq As Single)
  Dim t As Long, qqq As Single
  Speed = Speed + 1
  Do
    qq = Timer
    qqq = qq
    SetColor Picture1, Expl
    For t = 1 To zWidth / 20
      DrawCircle Picture1, (zWidth / 20), (zWidth / 20), t
      If t Mod Speed = 0 Then Picture1.Refresh
    Next t
    Picture1.Refresh
    SetColor Picture1, Sky
    For t = 1 To zWidth / 20
      DrawCircle Picture1, (zWidth / 20), (zWidth / 20), t
      If t Mod Speed = 0 Then Picture1.Refresh
    Next t
    Picture1.Refresh
    qq = Timer - qq
    If qq > SPD Then Speed = Speed + 1
  Loop Until (qq <= SPD) Or (Speed >= (zWidth / 20))
  If ((qqq - SPD) < (SPD - qq)) And (Speed > 1) Then Speed = Speed - 1
End Sub

Private Sub SetSpeed()
  Dim t As Long, qq As Single
  qq = Timer
  Speed = 8
  SetColor Picture1, Expl, GWIDTH
  For t = 1 To zWidth / 20
    DrawCircle Picture1, (zWidth / 20), (zWidth / 20), t
    If t Mod Speed = 0 Then Picture1.Refresh
  Next t
  Picture1.Refresh
  SetColor Picture1, Sky
  For t = 1 To zWidth / 20
    DrawCircle Picture1, (zWidth / 20), (zWidth / 20), t
    If t Mod Speed = 0 Then Picture1.Refresh
  Next t
  Picture1.Refresh
  qq = Timer - qq
  If qq < SPD Then
    Slower qq
  ElseIf qq > SPD Then
    Faster qq
  End If
End Sub
