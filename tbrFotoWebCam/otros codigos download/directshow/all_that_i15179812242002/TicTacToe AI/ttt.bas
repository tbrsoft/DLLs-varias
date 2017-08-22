Attribute VB_Name = "Module1"
Option Explicit

'Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global a, b, c, d, d1, d2, e, f, g, h As Integer
Global C0, C1, C2, C3, C4, C5, C6, C7, C8 As Object
Global o, x As String * 1, snd As String

Public Sub check()
  If C0.Caption = "X" And C1.Caption = "X" And C2.Caption = "X" Then d = 1
  If C3.Caption = "X" And C4.Caption = "X" And C5.Caption = "X" Then d = 1
  If C6.Caption = "X" And C7.Caption = "X" And C8.Caption = "X" Then d = 1
  If C0.Caption = "X" And C3.Caption = "X" And C6.Caption = "X" Then d = 1
  If C1.Caption = "X" And C4.Caption = "X" And C7.Caption = "X" Then d = 1
  If C2.Caption = "X" And C5.Caption = "X" And C8.Caption = "X" Then d = 1
  If C0.Caption = "X" And C4.Caption = "X" And C8.Caption = "X" Then d = 1
  If C2.Caption = "X" And C4.Caption = "X" And C6.Caption = "X" Then d = 1
  If C0.Caption = "O" And C1.Caption = "O" And C2.Caption = "O" Then d = 2
  If C3.Caption = "O" And C4.Caption = "O" And C5.Caption = "O" Then d = 2
  If C6.Caption = "O" And C7.Caption = "O" And C8.Caption = "O" Then d = 2
  If C0.Caption = "O" And C3.Caption = "O" And C6.Caption = "O" Then d = 2
  If C1.Caption = "O" And C4.Caption = "O" And C7.Caption = "O" Then d = 2
  If C2.Caption = "O" And C5.Caption = "O" And C8.Caption = "O" Then d = 2
  If C0.Caption = "O" And C4.Caption = "O" And C8.Caption = "O" Then d = 2
  If C2.Caption = "O" And C4.Caption = "O" And C6.Caption = "O" Then d = 2
  If d = 1 Then
    d1 = d1 + 1
    Form1.sc1.Caption = d1
    MsgBox "Player 1 win !", 64, "Game Over"
  End If
  If d = 2 Then
    d2 = d2 + 1
    Form1.sc2.Caption = d2
    MsgBox "Player 2 win !", 64, "Game Over"
  End If
  If c = 9 And d = 0 And f <> 0 Then MsgBox "Draw !", 64, "Game Over"
  If c = 9 Or d <> 0 Then
    Module1.restart
    Exit Sub
  End If
  a = -a
  If a = 1 Then
    Form1.Caption = "Game #" & Str$(g) & ", Player # 1"
  Else
    Form1.Caption = "Game #" & Str$(g) & ", Player # 2"
  End If
End Sub

Public Sub Main()
  Randomize Timer
  Set C0 = Form1.Label(0): Set C1 = Form1.Label(1): Set C2 = Form1.Label(2)
  Set C3 = Form1.Label(3): Set C4 = Form1.Label(4): Set C5 = Form1.Label(5)
  Set C6 = Form1.Label(6): Set C7 = Form1.Label(7): Set C8 = Form1.Label(8)
  a = 1: c = 0: d = 0: d1 = 0: d2 = 0: e = 1: f = 1: g = 1: h = 2
  If Right$(App.Path, 1) = "\" Then
    snd = App.Path & "done.wav"
  Else
    snd = App.Path & "\done.wav"
  End If
  Form1.Show
End Sub

Public Sub restart()
  Dim t As Long
  Form1.Timer1.Enabled = False
  c = 0: d = 0: e = -e
  If e = 1 Then
    a = 1
  Else
    a = -1
  End If
  For t = 0 To 8
    Form1.Label(t).Caption = ""
  Next t
  Form1.List1.Clear
  g = g + 1
  If a = 1 Then
    Form1.Caption = "Game #" & Str$(g) & ", Player # 1"
  Else
    Form1.Caption = "Game #" & Str$(g) & ", Player # 2"
  End If
  Form1.Refresh
  Form1.Timer1.Enabled = True
End Sub

Public Sub h1()
If h > 0 Then
  If C0.Caption = x And C1.Caption = x And C2.Caption = "" Then b = 2
  If C3.Caption = x And C4.Caption = x And C5.Caption = "" Then b = 5
  If C6.Caption = x And C7.Caption = x And C8.Caption = "" Then b = 8
  If C0.Caption = x And C2.Caption = x And C1.Caption = "" Then b = 1
  If C3.Caption = x And C5.Caption = x And C4.Caption = "" Then b = 4
  If C6.Caption = x And C8.Caption = x And C7.Caption = "" Then b = 7
  If C1.Caption = x And C2.Caption = x And C0.Caption = "" Then b = 0
  If C4.Caption = x And C5.Caption = x And C3.Caption = "" Then b = 3
  If C7.Caption = x And C8.Caption = x And C6.Caption = "" Then b = 6
  If C0.Caption = x And C3.Caption = x And C6.Caption = "" Then b = 6
  If C1.Caption = x And C4.Caption = x And C7.Caption = "" Then b = 7
  If C2.Caption = x And C5.Caption = x And C8.Caption = "" Then b = 8
  If C0.Caption = x And C6.Caption = x And C3.Caption = "" Then b = 3
  If C1.Caption = x And C7.Caption = x And C4.Caption = "" Then b = 4
  If C2.Caption = x And C8.Caption = x And C5.Caption = "" Then b = 5
  If C3.Caption = x And C6.Caption = x And C0.Caption = "" Then b = 0
  If C4.Caption = x And C7.Caption = x And C1.Caption = "" Then b = 1
  If C5.Caption = x And C8.Caption = x And C2.Caption = "" Then b = 2
  If C0.Caption = x And C4.Caption = x And C8.Caption = "" Then b = 8
  If C0.Caption = x And C8.Caption = x And C4.Caption = "" Then b = 4
  If C4.Caption = x And C8.Caption = x And C0.Caption = "" Then b = 0
  If C2.Caption = x And C4.Caption = x And C6.Caption = "" Then b = 6
  If C2.Caption = x And C6.Caption = x And C4.Caption = "" Then b = 4
  If C4.Caption = x And C6.Caption = x And C2.Caption = "" Then b = 2
  If C0.Caption = o And C1.Caption = o And C2.Caption = "" Then b = 2
  If C3.Caption = o And C4.Caption = o And C5.Caption = "" Then b = 5
  If C6.Caption = o And C7.Caption = o And C8.Caption = "" Then b = 8
  If C0.Caption = o And C2.Caption = o And C1.Caption = "" Then b = 1
  If C3.Caption = o And C5.Caption = o And C4.Caption = "" Then b = 4
  If C6.Caption = o And C8.Caption = o And C7.Caption = "" Then b = 7
  If C1.Caption = o And C2.Caption = o And C0.Caption = "" Then b = 0
  If C4.Caption = o And C5.Caption = o And C3.Caption = "" Then b = 3
  If C7.Caption = o And C8.Caption = o And C6.Caption = "" Then b = 6
  If C0.Caption = o And C3.Caption = o And C6.Caption = "" Then b = 6
  If C1.Caption = o And C4.Caption = o And C7.Caption = "" Then b = 7
  If C2.Caption = o And C5.Caption = o And C8.Caption = "" Then b = 8
  If C0.Caption = o And C6.Caption = o And C3.Caption = "" Then b = 3
  If C1.Caption = o And C7.Caption = o And C4.Caption = "" Then b = 4
  If C2.Caption = o And C8.Caption = o And C5.Caption = "" Then b = 5
  If C3.Caption = o And C6.Caption = o And C0.Caption = "" Then b = 0
  If C4.Caption = o And C7.Caption = o And C1.Caption = "" Then b = 1
  If C5.Caption = o And C8.Caption = o And C2.Caption = "" Then b = 2
  If C0.Caption = o And C4.Caption = o And C8.Caption = "" Then b = 8
  If C0.Caption = o And C8.Caption = o And C4.Caption = "" Then b = 4
  If C4.Caption = o And C8.Caption = o And C0.Caption = "" Then b = 0
  If C2.Caption = o And C4.Caption = o And C6.Caption = "" Then b = 6
  If C2.Caption = o And C6.Caption = o And C4.Caption = "" Then b = 4
  If C4.Caption = o And C6.Caption = o And C2.Caption = "" Then b = 2
End If
End Sub

Public Sub play()
Dim z As Integer
DoEvents
If f = 1 Then Sleep 250
debut:
b = Int(Rnd * 9)
If Form1.Label(b).Caption <> "" Then GoTo debut
z = Int(Rnd * (10 - c))
If h > 0 And z = 0 And C4.Caption = "" Then b = 4
If h > 1 Then
  If c = 1 And C4.Caption = x And (b = 1 Or b = 3 Or b = 5 Or b = 7) Then GoTo debut
  If c = 2 Then
    If C4.Caption = x Then
      If C1.Caption = o And b = 7 Then GoTo debut
      If C3.Caption = o And b = 5 Then GoTo debut
      If C7.Caption = o And b = 1 Then GoTo debut
      If C5.Caption = o And b = 3 Then GoTo debut
    End If
    If C1.Caption = x And C7.Caption = o And b <> 6 And b <> 8 Then GoTo debut
    If C3.Caption = x And C5.Caption = o And b <> 2 And b <> 8 Then GoTo debut
    If C7.Caption = x And C1.Caption = o And b <> 0 And b <> 2 Then GoTo debut
    If C5.Caption = x And C3.Caption = o And b <> 0 And b <> 6 Then GoTo debut
  End If
  If c = 3 Then
    If C1.Caption = x And C3.Caption = x And C5.Caption = o And b <> 0 And b <> 6 Then GoTo debut
    If C3.Caption = x And C7.Caption = x And C1.Caption = o And b <> 6 And b <> 8 Then GoTo debut
    If C7.Caption = x And C5.Caption = x And C3.Caption = o And b <> 2 And b <> 8 Then GoTo debut
    If C5.Caption = x And C1.Caption = x And C7.Caption = o And b <> 0 And b <> 2 Then GoTo debut
    If C1.Caption = x And C3.Caption = x And C7.Caption = o And b <> 0 And b <> 2 Then GoTo debut
    If C3.Caption = x And C7.Caption = x And C5.Caption = o And b <> 0 And b <> 6 Then GoTo debut
    If C7.Caption = x And C5.Caption = x And C1.Caption = o And b <> 6 And b <> 8 Then GoTo debut
    If C5.Caption = x And C1.Caption = x And C3.Caption = o And b <> 2 And b <> 8 Then GoTo debut
    
    If C8.Caption = o And C5.Caption = x And C7.Caption = x And (b = 0 Or b = 2 Or b = 6) Then GoTo debut
    If C2.Caption = o And C1.Caption = x And C5.Caption = x And (b = 0 Or b = 6 Or b = 8) Then GoTo debut
    If C0.Caption = o And C1.Caption = x And C3.Caption = x And (b = 2 Or b = 6 Or b = 8) Then GoTo debut
    If C6.Caption = o And C3.Caption = x And C7.Caption = x And (b = 0 Or b = 2 Or b = 8) Then GoTo debut
  End If
  If c = 4 Then
    If C4.Caption = o Then
      If C1.Caption = o And C5.Caption = x And C7.Caption = x And b <> 0 And b <> 2 Then GoTo debut
      If C3.Caption = o And C1.Caption = x And C5.Caption = x And b <> 0 And b <> 6 Then GoTo debut
      If C7.Caption = o And C3.Caption = x And C1.Caption = x And b <> 6 And b <> 8 Then GoTo debut
      If C5.Caption = o And C7.Caption = x And C3.Caption = x And b <> 2 And b <> 8 Then GoTo debut
      If C1.Caption = o And C3.Caption = x And C7.Caption = x And b <> 0 And b <> 2 Then GoTo debut
      If C3.Caption = o And C7.Caption = x And C5.Caption = x And b <> 0 And b <> 6 Then GoTo debut
      If C7.Caption = o And C5.Caption = x And C1.Caption = x And b <> 6 And b <> 8 Then GoTo debut
      If C5.Caption = o And C1.Caption = x And C3.Caption = x And b <> 2 And b <> 8 Then GoTo debut
    End If
    If C1.Caption = o And C5.Caption = o And C2.Caption = x And C3.Caption = x And b = 8 Then GoTo debut
    If C3.Caption = o And C1.Caption = o And C0.Caption = x And C7.Caption = x And b = 2 Then GoTo debut
    If C7.Caption = o And C3.Caption = o And C6.Caption = x And C5.Caption = x And b = 0 Then GoTo debut
    If C5.Caption = o And C7.Caption = o And C8.Caption = x And C1.Caption = x And b = 6 Then GoTo debut
    If C1.Caption = o And C3.Caption = o And C0.Caption = x And C5.Caption = x And b = 6 Then GoTo debut
    If C3.Caption = o And C7.Caption = o And C6.Caption = x And C1.Caption = x And b = 8 Then GoTo debut
    If C7.Caption = o And C5.Caption = o And C8.Caption = x And C3.Caption = x And b = 2 Then GoTo debut
    If C5.Caption = o And C1.Caption = o And C2.Caption = x And C7.Caption = x And b = 0 Then GoTo debut
  End If
  If c = 5 Then
    If C0.Caption = x And C2.Caption = x And C7.Caption = x And C1.Caption = o And C6.Caption = o And b = 3 Then GoTo debut
    If C6.Caption = x And C0.Caption = x And C5.Caption = x And C3.Caption = o And C8.Caption = o And b = 7 Then GoTo debut
    If C8.Caption = x And C6.Caption = x And C1.Caption = x And C7.Caption = o And C2.Caption = o And b = 5 Then GoTo debut
    If C2.Caption = x And C8.Caption = x And C3.Caption = x And C5.Caption = o And C0.Caption = o And b = 1 Then GoTo debut
    If C0.Caption = x And C2.Caption = x And C7.Caption = x And C1.Caption = o And C8.Caption = o And b = 5 Then GoTo debut
    If C6.Caption = x And C0.Caption = x And C5.Caption = x And C3.Caption = o And C2.Caption = o And b = 1 Then GoTo debut
    If C8.Caption = x And C6.Caption = x And C1.Caption = x And C7.Caption = o And C0.Caption = o And b = 3 Then GoTo debut
    If C2.Caption = x And C8.Caption = x And C3.Caption = x And C5.Caption = o And C6.Caption = o And b = 7 Then GoTo debut
  End If
End If
Module1.h2
If h > 2 Then
  If c = 2 Then
    If C2.Caption = x And C6.Caption = o And b <> 0 And b <> 8 Then GoTo debut
    If C0.Caption = x And C8.Caption = o And b <> 2 And b <> 6 Then GoTo debut
    If C6.Caption = x And C2.Caption = o And b <> 0 And b <> 8 Then GoTo debut
    If C8.Caption = x And C0.Caption = o And b <> 2 And b <> 6 Then GoTo debut
  End If
  If c = 3 Then
    If C4.Caption = x Then
      If C1.Caption = x And C7.Caption = o And (b = 3 Or b = 5) Then GoTo debut
      If C3.Caption = x And C5.Caption = o And (b = 1 Or b = 7) Then GoTo debut
      If C7.Caption = x And C1.Caption = o And (b = 3 Or b = 5) Then GoTo debut
      If C5.Caption = x And C3.Caption = o And (b = 1 Or b = 7) Then GoTo debut
    End If
    If C4.Caption = o Then
      If C0.Caption = x And C8.Caption = x And (b = 2 Or b = 6) Then GoTo debut
      If C2.Caption = x And C6.Caption = x And (b = 0 Or b = 8) Then GoTo debut
      
      If C0.Caption = x And C7.Caption = x And (b = 1 Or b = 2) Then GoTo debut
      If C5.Caption = x And C6.Caption = x And (b = 0 Or b = 3) Then GoTo debut
      If C1.Caption = x And C8.Caption = x And (b = 6 Or b = 7) Then GoTo debut
      If C2.Caption = x And C3.Caption = x And (b = 5 Or b = 8) Then GoTo debut
      If C2.Caption = x And C7.Caption = x And (b = 0 Or b = 1) Then GoTo debut
      If C0.Caption = x And C5.Caption = x And (b = 3 Or b = 6) Then GoTo debut
      If C1.Caption = x And C6.Caption = x And (b = 7 Or b = 8) Then GoTo debut
      If C3.Caption = x And C8.Caption = x And (b = 2 Or b = 5) Then GoTo debut

      If C1.Caption = x And C3.Caption = x And (b = 5 Or b = 7 Or b = 8) Then GoTo debut
      If C3.Caption = x And C7.Caption = x And (b = 1 Or b = 2 Or b = 5) Then GoTo debut
      If C5.Caption = x And C7.Caption = x And (b = 0 Or b = 1 Or b = 3) Then GoTo debut
      If C1.Caption = x And C5.Caption = x And (b = 3 Or b = 6 Or b = 7) Then GoTo debut
    End If
    If b <> 4 Then
      If C0.Caption = o And C1.Caption = x And C6.Caption = x And b <> 7 Then GoTo debut
      If C6.Caption = o And C3.Caption = x And C8.Caption = x And b <> 5 Then GoTo debut
      If C8.Caption = o And C7.Caption = x And C2.Caption = x And b <> 1 Then GoTo debut
      If C2.Caption = o And C5.Caption = x And C0.Caption = x And b <> 3 Then GoTo debut
      If C2.Caption = o And C1.Caption = x And C8.Caption = x And b <> 7 Then GoTo debut
      If C0.Caption = o And C3.Caption = x And C2.Caption = x And b <> 5 Then GoTo debut
      If C6.Caption = o And C7.Caption = x And C0.Caption = x And b <> 1 Then GoTo debut
      If C8.Caption = o And C5.Caption = x And C6.Caption = x And b <> 3 Then GoTo debut
    End If
  End If
  If c = 4 Then
    If C4.Caption = x Then
      If C0.Caption = x And C1.Caption = o And C8.Caption = o And b = 7 Then GoTo debut
      If C6.Caption = x And C2.Caption = o And C3.Caption = o And b = 5 Then GoTo debut
      If C8.Caption = x And C0.Caption = o And C7.Caption = o And b = 1 Then GoTo debut
      If C2.Caption = x And C5.Caption = o And C6.Caption = o And b = 3 Then GoTo debut
      If C2.Caption = x And C1.Caption = o And C6.Caption = o And b = 7 Then GoTo debut
      If C0.Caption = x And C3.Caption = o And C8.Caption = o And b = 5 Then GoTo debut
      If C6.Caption = x And C2.Caption = o And C7.Caption = o And b = 1 Then GoTo debut
      If C8.Caption = x And C0.Caption = o And C5.Caption = o And b = 3 Then GoTo debut
    End If
    If C4.Caption = o Then
      If C0.Caption = x And C7.Caption = x And C1.Caption = o And b = 2 Then GoTo debut
      If C5.Caption = x And C6.Caption = x And C3.Caption = o And b = 0 Then GoTo debut
      If C1.Caption = x And C8.Caption = x And C7.Caption = o And b = 6 Then GoTo debut
      If C2.Caption = x And C3.Caption = x And C5.Caption = o And b = 8 Then GoTo debut
      If C2.Caption = x And C7.Caption = x And C1.Caption = o And b = 0 Then GoTo debut
      If C0.Caption = x And C5.Caption = x And C3.Caption = o And b = 6 Then GoTo debut
      If C1.Caption = x And C6.Caption = x And C7.Caption = o And b = 8 Then GoTo debut
      If C3.Caption = x And C8.Caption = x And C5.Caption = o And b = 2 Then GoTo debut
    End If
    If b = 4 Then
      If C2.Caption = o And C5.Caption = o And C3.Caption = x And C8.Caption = x Then GoTo debut
      If C0.Caption = o And C1.Caption = o And C2.Caption = x And C7.Caption = x Then GoTo debut
      If C3.Caption = o And C6.Caption = o And C0.Caption = x And C5.Caption = x Then GoTo debut
      If C7.Caption = o And C8.Caption = o And C1.Caption = x And C6.Caption = x Then GoTo debut
      If C0.Caption = o And C3.Caption = o And C5.Caption = x And C6.Caption = x Then GoTo debut
      If C6.Caption = o And C7.Caption = o And C1.Caption = x And C8.Caption = x Then GoTo debut
      If C5.Caption = o And C8.Caption = o And C2.Caption = x And C3.Caption = x Then GoTo debut
      If C1.Caption = o And C2.Caption = o And C0.Caption = x And C7.Caption = x Then GoTo debut
    End If
  End If
  If c = 5 Then
    If C4.Caption = o Then
      If C3.Caption = o And C1.Caption = x And C5.Caption = x And C7.Caption = x And b <> 0 And b <> 6 Then GoTo debut
      If C7.Caption = o And C1.Caption = x And C3.Caption = x And C5.Caption = x And b <> 6 And b <> 8 Then GoTo debut
      If C5.Caption = o And C1.Caption = x And C3.Caption = x And C7.Caption = x And b <> 2 And b <> 8 Then GoTo debut
      If C1.Caption = o And C3.Caption = x And C5.Caption = x And C7.Caption = x And b <> 0 And b <> 2 Then GoTo debut

      If C3.Caption = o And C0.Caption = x And C5.Caption = x And C7.Caption = x And b <> 2 And b <> 8 Then GoTo debut
      If C7.Caption = o And C1.Caption = x And C5.Caption = x And C6.Caption = x And b <> 0 And b <> 2 Then GoTo debut
      If C5.Caption = o And C1.Caption = x And C3.Caption = x And C8.Caption = x And b <> 0 And b <> 6 Then GoTo debut
      If C1.Caption = o And C2.Caption = x And C3.Caption = x And C7.Caption = x And b <> 6 And b <> 8 Then GoTo debut
      If C5.Caption = o And C2.Caption = x And C3.Caption = x And C7.Caption = x And b <> 0 And b <> 6 Then GoTo debut
      If C1.Caption = o And C0.Caption = x And C5.Caption = x And C7.Caption = x And b <> 6 And b <> 8 Then GoTo debut
      If C3.Caption = o And C1.Caption = x And C5.Caption = x And C6.Caption = x And b <> 2 And b <> 8 Then GoTo debut
      If C7.Caption = o And C1.Caption = x And C3.Caption = x And C8.Caption = x And b <> 0 And b <> 2 Then GoTo debut
    End If
    If b <> 4 Then
      If C1.Caption = x And C3.Caption = x And C8.Caption = x And C5.Caption = o And C6.Caption = o And b <> 0 Then GoTo debut
      If C3.Caption = x And C7.Caption = x And C2.Caption = x And C1.Caption = o And C8.Caption = o And b <> 6 Then GoTo debut
      If C7.Caption = x And C5.Caption = x And C0.Caption = x And C3.Caption = o And C2.Caption = o And b <> 8 Then GoTo debut
      If C5.Caption = x And C1.Caption = x And C6.Caption = x And C7.Caption = o And C0.Caption = o And b <> 2 Then GoTo debut
      If C1.Caption = x And C3.Caption = x And C8.Caption = x And C2.Caption = o And C7.Caption = o And b <> 0 Then GoTo debut
      If C3.Caption = x And C7.Caption = x And C2.Caption = x And C0.Caption = o And C5.Caption = o And b <> 6 Then GoTo debut
      If C7.Caption = x And C5.Caption = x And C0.Caption = x And C6.Caption = o And C1.Caption = o And b <> 8 Then GoTo debut
      If C5.Caption = x And C1.Caption = x And C6.Caption = x And C8.Caption = o And C3.Caption = o And b <> 2 Then GoTo debut
    End If
  End If
End If
Module1.h3
If h > 3 Then
  If c = 1 Then
    If C1.Caption = x And (b = 3 Or b = 5 Or b = 6 Or b = 8 Or b = 7) Then GoTo debut
    If C3.Caption = x And (b = 1 Or b = 7 Or b = 2 Or b = 8 Or b = 5) Then GoTo debut
    If C7.Caption = x And (b = 3 Or b = 5 Or b = 0 Or b = 2 Or b = 1) Then GoTo debut
    If C5.Caption = x And (b = 1 Or b = 7 Or b = 0 Or b = 6 Or b = 3) Then GoTo debut
  End If
  If c = 2 Then
    If C3.Caption = o And C6.Caption = x And b = 0 Then GoTo debut
    If C7.Caption = o And C8.Caption = x And b = 6 Then GoTo debut
    If C5.Caption = o And C2.Caption = x And b = 8 Then GoTo debut
    If C1.Caption = o And C0.Caption = x And b = 2 Then GoTo debut
    If C5.Caption = o And C8.Caption = x And b = 2 Then GoTo debut
    If C1.Caption = o And C2.Caption = x And b = 0 Then GoTo debut
    If C3.Caption = o And C0.Caption = x And b = 6 Then GoTo debut
    If C7.Caption = o And C6.Caption = x And b = 8 Then GoTo debut
    
    If C3.Caption = o And C6.Caption = x And (b = 1 Or b = 5) Then GoTo debut
    If C7.Caption = o And C8.Caption = x And (b = 1 Or b = 3) Then GoTo debut
    If C5.Caption = o And C2.Caption = x And (b = 3 Or b = 7) Then GoTo debut
    If C1.Caption = o And C0.Caption = x And (b = 5 Or b = 7) Then GoTo debut
    If C5.Caption = o And C8.Caption = x And (b = 1 Or b = 3) Then GoTo debut
    If C1.Caption = o And C2.Caption = x And (b = 3 Or b = 7) Then GoTo debut
    If C3.Caption = o And C0.Caption = x And (b = 5 Or b = 7) Then GoTo debut
    If C7.Caption = o And C6.Caption = x And (b = 1 Or b = 5) Then GoTo debut
  End If
  If c = 3 Then
    If C4.Caption = o Then
      If C1.Caption = x And C7.Caption = x And b <> 3 And b <> 5 Then GoTo debut
      If C3.Caption = x And C5.Caption = x And b <> 1 And b <> 7 Then GoTo debut
    End If
    If C4.Caption = x Then
      If C0.Caption = x And C8.Caption = o And b <> 2 And b <> 6 Then GoTo debut
      If C6.Caption = x And C2.Caption = o And b <> 0 And b <> 8 Then GoTo debut
      If C8.Caption = x And C0.Caption = o And b <> 2 And b <> 6 Then GoTo debut
      If C2.Caption = x And C6.Caption = o And b <> 0 And b <> 8 Then GoTo debut
    End If
    If C0.Caption = x And C3.Caption = x And C6.Caption = o And b = 2 Then GoTo debut
    If C6.Caption = x And C7.Caption = x And C8.Caption = o And b = 0 Then GoTo debut
    If C8.Caption = x And C5.Caption = x And C2.Caption = o And b = 6 Then GoTo debut
    If C2.Caption = x And C1.Caption = x And C0.Caption = o And b = 8 Then GoTo debut
    If C2.Caption = x And C5.Caption = x And C8.Caption = o And b = 0 Then GoTo debut
    If C0.Caption = x And C1.Caption = x And C2.Caption = o And b = 6 Then GoTo debut
    If C6.Caption = x And C3.Caption = x And C0.Caption = o And b = 8 Then GoTo debut
    If C8.Caption = x And C7.Caption = x And C6.Caption = o And b = 2 Then GoTo debut
  End If
  If c = 4 And C4.Caption = o Then
    If C8.Caption = o And C0.Caption = x And C5.Caption = x And b <> 6 And b <> 7 Then GoTo debut
    If C2.Caption = o And C1.Caption = x And C6.Caption = x And b <> 5 And b <> 8 Then GoTo debut
    If C0.Caption = o And C3.Caption = x And C8.Caption = x And b <> 1 And b <> 2 Then GoTo debut
    If C6.Caption = o And C2.Caption = x And C7.Caption = x And b <> 0 And b <> 3 Then GoTo debut
    If C8.Caption = o And C0.Caption = x And C7.Caption = x And b <> 2 And b <> 5 Then GoTo debut
    If C2.Caption = o And C5.Caption = x And C6.Caption = x And b <> 0 And b <> 1 Then GoTo debut
    If C0.Caption = o And C1.Caption = x And C8.Caption = x And b <> 3 And b <> 6 Then GoTo debut
    If C6.Caption = o And C2.Caption = x And C2.Caption = x And b <> 7 And b <> 8 Then GoTo debut
  End If
  If c = 5 Then
    If C1.Caption = x And C5.Caption = x And C6.Caption = x And C2.Caption = o And C7.Caption = o And b = 8 Then GoTo debut
    If C3.Caption = x And C1.Caption = x And C8.Caption = x And C0.Caption = o And C5.Caption = o And b = 2 Then GoTo debut
    If C7.Caption = x And C3.Caption = x And C2.Caption = x And C6.Caption = o And C1.Caption = o And b = 0 Then GoTo debut
    If C5.Caption = x And C7.Caption = x And C0.Caption = x And C8.Caption = o And C3.Caption = o And b = 6 Then GoTo debut
    If C1.Caption = x And C3.Caption = x And C8.Caption = x And C0.Caption = o And C7.Caption = o And b = 6 Then GoTo debut
    If C3.Caption = x And C7.Caption = x And C2.Caption = x And C6.Caption = o And C5.Caption = o And b = 8 Then GoTo debut
    If C7.Caption = x And C1.Caption = x And C0.Caption = x And C8.Caption = o And C1.Caption = o And b = 2 Then GoTo debut
    If C5.Caption = x And C5.Caption = x And C6.Caption = x And C2.Caption = o And C3.Caption = o And b = 0 Then GoTo debut
        
    If C3.Caption = x And C7.Caption = x And C8.Caption = x And C5.Caption = o And C6.Caption = o And b = 2 Then GoTo debut
    If C7.Caption = x And C5.Caption = x And C2.Caption = x And C1.Caption = o And C8.Caption = o And b = 0 Then GoTo debut
    If C5.Caption = x And C1.Caption = x And C0.Caption = x And C3.Caption = o And C2.Caption = o And b = 6 Then GoTo debut
    If C1.Caption = x And C3.Caption = x And C6.Caption = x And C7.Caption = o And C0.Caption = o And b = 8 Then GoTo debut
    If C5.Caption = x And C6.Caption = x And C7.Caption = x And C3.Caption = o And C8.Caption = o And b = 0 Then GoTo debut
    If C1.Caption = x And C8.Caption = x And C5.Caption = x And C7.Caption = o And C2.Caption = o And b = 6 Then GoTo debut
    If C3.Caption = x And C2.Caption = x And C1.Caption = x And C5.Caption = o And C0.Caption = o And b = 8 Then GoTo debut
    If C7.Caption = x And C0.Caption = x And C3.Caption = x And C1.Caption = o And C6.Caption = o And b = 2 Then GoTo debut
  End If
End If
Module1.h4
If h > 3 Then
  If c = 5 Then
    If C0.Caption = x And C1.Caption = o And C5.Caption = x And C7.Caption = x And C8.Caption = o And b = 2 Then GoTo debut
    If C2.Caption = x And C3.Caption = x And C5.Caption = o And C6.Caption = o And C7.Caption = x And b = 8 Then GoTo debut
    If C0.Caption = o And C1.Caption = x And C3.Caption = x And C7.Caption = o And C8.Caption = x And b = 6 Then GoTo debut
    If C1.Caption = x And C2.Caption = o And C3.Caption = o And C5.Caption = x And C6.Caption = x And b = 0 Then GoTo debut
    If C1.Caption = o And C2.Caption = x And C3.Caption = x And C6.Caption = o And C7.Caption = x And b = 0 Then GoTo debut
    If C0.Caption = o And C1.Caption = x And C3.Caption = x And C5.Caption = o And C8.Caption = x And b = 2 Then GoTo debut
    If C1.Caption = x And C2.Caption = o And C5.Caption = x And C6.Caption = x And C7.Caption = o And b = 8 Then GoTo debut
    If C0.Caption = x And C3.Caption = o And C5.Caption = x And C7.Caption = x And C8.Caption = o And b = 6 Then GoTo debut
  End If
End If
Module1.h5
Module1.h1
Form1.List1.AddItem b
Form1.Label(b).Caption = o
If f <> 0 Then sndPlaySound snd, 1
c = c + 1
Module1.check
End Sub

Public Sub h2()
If h > 1 Then
  If c = 3 Then
    If C1.Caption = x And C3.Caption = x And C2.Caption = o Then b = 8
    If C3.Caption = x And C7.Caption = x And C0.Caption = o Then b = 2
    If C7.Caption = x And C5.Caption = x And C6.Caption = o Then b = 0
    If C5.Caption = x And C1.Caption = x And C8.Caption = o Then b = 6
    If C1.Caption = x And C5.Caption = x And C0.Caption = o Then b = 6
    If C3.Caption = x And C1.Caption = x And C6.Caption = o Then b = 8
    If C7.Caption = x And C3.Caption = x And C8.Caption = o Then b = 2
    If C5.Caption = x And C7.Caption = x And C2.Caption = o Then b = 0
  End If
  If c = 4 Then
    If (C1.Caption = o And C8.Caption = o And C2.Caption = x) _
      And (C0.Caption = x Or C3.Caption = x Or C5.Caption = x) Then b = 7
    If (C2.Caption = o And C3.Caption = o And C0.Caption = x) _
      And (C1.Caption = x Or C6.Caption = x Or C7.Caption = x) Then b = 5
    If (C0.Caption = o And C7.Caption = o And C6.Caption = x) _
      And (C3.Caption = x Or C5.Caption = x Or C8.Caption = x) Then b = 1
    If (C5.Caption = o And C6.Caption = o And C8.Caption = x) _
      And (C1.Caption = x Or C2.Caption = x Or C7.Caption = x) Then b = 3
    If (C1.Caption = o And C6.Caption = o And C0.Caption = x) _
      And (C2.Caption = x Or C3.Caption = x Or C5.Caption = x) Then b = 7
    If (C3.Caption = o And C8.Caption = o And C6.Caption = x) _
      And (C1.Caption = x Or C0.Caption = x Or C7.Caption = x) Then b = 5
    If (C7.Caption = o And C2.Caption = o And C8.Caption = x) _
      And (C3.Caption = x Or C5.Caption = x Or C6.Caption = x) Then b = 1
    If (C5.Caption = o And C0.Caption = o And C2.Caption = x) _
      And (C1.Caption = x Or C8.Caption = x Or C7.Caption = x) Then b = 3
        
    If C0.Caption = o And C6.Caption = o And C2.Caption = x And C3.Caption = x Then b = 8
    If C6.Caption = o And C8.Caption = o And C0.Caption = x And C7.Caption = x Then b = 2
    If C8.Caption = o And C2.Caption = o And C6.Caption = x And C5.Caption = x Then b = 0
    If C2.Caption = o And C0.Caption = o And C8.Caption = x And C1.Caption = x Then b = 6
    If C6.Caption = o And C8.Caption = o And C2.Caption = x And C7.Caption = x Then b = 0
    If C8.Caption = o And C2.Caption = o And C0.Caption = x And C5.Caption = x Then b = 6
    If C2.Caption = o And C0.Caption = o And C6.Caption = x And C1.Caption = x Then b = 2
    If C0.Caption = o And C6.Caption = o And C8.Caption = x And C3.Caption = x Then b = 8
  End If
  If c = 5 Then
    If C1.Caption = x And C5.Caption = x And C8.Caption = x And C2.Caption = o And C3.Caption = o Then b = 6
    If C3.Caption = x And C1.Caption = x And C2.Caption = x And C0.Caption = o And C7.Caption = o Then b = 8
    If C7.Caption = x And C3.Caption = x And C0.Caption = x And C6.Caption = o And C5.Caption = o Then b = 2
    If C5.Caption = x And C7.Caption = x And C6.Caption = x And C8.Caption = o And C1.Caption = o Then b = 0
    If C1.Caption = x And C3.Caption = x And C6.Caption = x And C0.Caption = o And C5.Caption = o Then b = 6
    If C3.Caption = x And C7.Caption = x And C8.Caption = x And C6.Caption = o And C1.Caption = o Then b = 8
    If C7.Caption = x And C5.Caption = x And C2.Caption = x And C8.Caption = o And C3.Caption = o Then b = 2
    If C5.Caption = x And C1.Caption = x And C0.Caption = x And C2.Caption = o And C7.Caption = o Then b = 0
  End If
  If c = 6 Then
    If C2.Caption = x And C7.Caption = x And C8.Caption = x And C1.Caption = o And C5.Caption = o And C6.Caption = o Then b = 3
    If C0.Caption = x And C5.Caption = x And C2.Caption = x And C3.Caption = o And C1.Caption = o And C8.Caption = o Then b = 7
    If C6.Caption = x And C1.Caption = x And C0.Caption = x And C7.Caption = o And C3.Caption = o And C2.Caption = o Then b = 5
    If C8.Caption = x And C3.Caption = x And C6.Caption = x And C5.Caption = o And C7.Caption = o And C0.Caption = o Then b = 1
    If C0.Caption = x And C6.Caption = x And C7.Caption = x And C1.Caption = o And C3.Caption = o And C8.Caption = o Then b = 5
    If C6.Caption = x And C8.Caption = x And C5.Caption = x And C3.Caption = o And C7.Caption = o And C2.Caption = o Then b = 1
    If C8.Caption = x And C2.Caption = x And C1.Caption = x And C7.Caption = o And C5.Caption = o And C0.Caption = o Then b = 3
    If C2.Caption = x And C0.Caption = x And C3.Caption = x And C5.Caption = o And C1.Caption = o And C6.Caption = o Then b = 7
                    
    If C0.Caption = o And C2.Caption = o And C7.Caption = o And C1.Caption = x And C3.Caption = x And C6.Caption = x Then b = 8
    If C6.Caption = o And C0.Caption = o And C5.Caption = o And C3.Caption = x And C7.Caption = x And C8.Caption = x Then b = 2
    If C8.Caption = o And C6.Caption = o And C1.Caption = o And C7.Caption = x And C5.Caption = x And C2.Caption = x Then b = 0
    If C2.Caption = o And C8.Caption = o And C3.Caption = o And C5.Caption = x And C1.Caption = x And C0.Caption = x Then b = 6
    If C0.Caption = o And C2.Caption = o And C7.Caption = o And C1.Caption = x And C5.Caption = x And C8.Caption = x Then b = 6
    If C6.Caption = o And C0.Caption = o And C5.Caption = o And C3.Caption = x And C1.Caption = x And C2.Caption = x Then b = 8
    If C8.Caption = o And C6.Caption = o And C1.Caption = o And C7.Caption = x And C3.Caption = x And C0.Caption = x Then b = 2
    If C2.Caption = o And C8.Caption = o And C3.Caption = o And C5.Caption = x And C7.Caption = x And C6.Caption = x Then b = 0
  End If
End If
End Sub

Public Sub h3()
If h > 2 Then
  If c = 1 Then
    If C6.Caption = x Then b = 4
    If C8.Caption = x Then b = 4
    If C2.Caption = x Then b = 4
    If C0.Caption = x Then b = 4
  End If
  If c = 2 Then
    If C4.Caption = o Then
      If C0.Caption = x Then b = 8
      If C6.Caption = x Then b = 2
      If C8.Caption = x Then b = 0
      If C2.Caption = x Then b = 6
    End If
    If C1.Caption = o And C5.Caption = x Then b = 4
    If C3.Caption = o And C1.Caption = x Then b = 4
    If C7.Caption = o And C3.Caption = x Then b = 4
    If C5.Caption = o And C7.Caption = x Then b = 4
    If C1.Caption = o And C3.Caption = x Then b = 4
    If C3.Caption = o And C7.Caption = x Then b = 4
    If C7.Caption = o And C5.Caption = x Then b = 4
    If C5.Caption = o And C1.Caption = x Then b = 4
  End If
  If c = 3 Then
    If (C2.Caption = o And C3.Caption = x) _
      And (C1.Caption = x Or C7.Caption = x) Then b = 8
    If (C0.Caption = o And C7.Caption = x) _
      And (C3.Caption = x Or C5.Caption = x) Then b = 2
    If (C6.Caption = o And C5.Caption = x) _
      And (C1.Caption = x Or C7.Caption = x) Then b = 0
    If (C8.Caption = o And C1.Caption = x) _
      And (C3.Caption = x Or C5.Caption = x) Then b = 6
    If (C0.Caption = o And C5.Caption = x) _
      And (C1.Caption = x Or C7.Caption = x) Then b = 6
    If (C6.Caption = o And C1.Caption = x) _
      And (C3.Caption = x Or C5.Caption = x) Then b = 8
    If (C8.Caption = o And C3.Caption = x) _
      And (C1.Caption = x Or C7.Caption = x) Then b = 2
    If (C2.Caption = o And C7.Caption = x) _
      And (C3.Caption = x Or C5.Caption = x) Then b = 0
  End If
  If c = 4 And C4.Caption = x Then
    If C1.Caption = x And C0.Caption = o And C7.Caption = o Then b = 6
    If C3.Caption = x And C5.Caption = o And C6.Caption = o Then b = 8
    If C7.Caption = x And C1.Caption = o And C8.Caption = o Then b = 2
    If C5.Caption = x And C2.Caption = o And C3.Caption = o Then b = 0
    If C1.Caption = x And C2.Caption = o And C7.Caption = o Then b = 8
    If C3.Caption = x And C0.Caption = o And C5.Caption = o Then b = 2
    If C7.Caption = x And C1.Caption = o And C6.Caption = o Then b = 0
    If C5.Caption = x And C3.Caption = o And C8.Caption = o Then b = 6

    If C7.Caption = x And C1.Caption = o And C3.Caption = o Then b = 0
    If C5.Caption = x And C3.Caption = o And C7.Caption = o Then b = 6
    If C1.Caption = x And C5.Caption = o And C7.Caption = o Then b = 8
    If C3.Caption = x And C1.Caption = o And C5.Caption = o Then b = 2
    If C7.Caption = x And C1.Caption = o And C5.Caption = o Then b = 2
    If C5.Caption = x And C1.Caption = o And C3.Caption = o Then b = 0
    If C1.Caption = x And C3.Caption = o And C7.Caption = o Then b = 6
    If C3.Caption = x And C5.Caption = o And C7.Caption = o Then b = 8
  End If
  If c = 6 And C4.Caption = x Then
    If C1.Caption = x And C3.Caption = x And C0.Caption = o And C5.Caption = o And C7.Caption = o Then b = 8
    If C3.Caption = x And C7.Caption = x And C1.Caption = o And C5.Caption = o And C6.Caption = o Then b = 2
    If C5.Caption = x And C7.Caption = x And C1.Caption = o And C3.Caption = o And C8.Caption = o Then b = 0
    If C1.Caption = x And C5.Caption = x And C2.Caption = o And C3.Caption = o And C7.Caption = o Then b = 6
  End If
End If
End Sub

Public Sub h4()
If h > 3 Then
  If c = 2 Then
    If C0.Caption = o And C2.Caption = x Then b = 3
    If C6.Caption = o And C0.Caption = x Then b = 7
    If C8.Caption = o And C6.Caption = x Then b = 5
    If C2.Caption = o And C8.Caption = x Then b = 1
    If C6.Caption = o And C8.Caption = x Then b = 3
    If C8.Caption = o And C2.Caption = x Then b = 7
    If C2.Caption = o And C0.Caption = x Then b = 5
    If C0.Caption = o And C6.Caption = x Then b = 1
       
    If C2.Caption = x And C3.Caption = o Then b = 0
    If C0.Caption = x And C7.Caption = o Then b = 6
    If C6.Caption = x And C5.Caption = o Then b = 8
    If C8.Caption = x And C1.Caption = o Then b = 2
    If C0.Caption = x And C5.Caption = o Then b = 2
    If C6.Caption = x And C1.Caption = o Then b = 0
    If C8.Caption = x And C3.Caption = o Then b = 6
    If C2.Caption = x And C7.Caption = o Then b = 8

    If (C0.Caption = o Or C2.Caption = o Or C6.Caption = o Or C8.Caption = o) _
      And (C1.Caption = x Or C3.Caption = x Or C5.Caption = x Or C7.Caption = x) Then b = 4
  End If
  If c = 3 Then
    If C1.Caption = x And C2.Caption = o And C6.Caption = x Then b = 4
    If C3.Caption = x And C0.Caption = o And C8.Caption = x Then b = 4
    If C7.Caption = x And C6.Caption = o And C2.Caption = x Then b = 4
    If C5.Caption = x And C8.Caption = o And C0.Caption = x Then b = 4
    If C1.Caption = x And C0.Caption = o And C8.Caption = x Then b = 4
    If C3.Caption = x And C6.Caption = o And C2.Caption = x Then b = 4
    If C7.Caption = x And C8.Caption = o And C0.Caption = x Then b = 4
    If C5.Caption = x And C2.Caption = o And C6.Caption = x Then b = 4
    
    If C0.Caption = x And C3.Caption = x And C6.Caption = o Then b = 7
    If C6.Caption = x And C7.Caption = x And C8.Caption = o Then b = 5
    If C5.Caption = x And C8.Caption = x And C2.Caption = o Then b = 1
    If C1.Caption = x And C2.Caption = x And C0.Caption = o Then b = 3
    If C2.Caption = x And C5.Caption = x And C8.Caption = o Then b = 7
    If C0.Caption = x And C1.Caption = x And C2.Caption = o Then b = 5
    If C3.Caption = x And C6.Caption = x And C0.Caption = o Then b = 1
    If C7.Caption = x And C8.Caption = x And C6.Caption = o Then b = 3
  End If
  If c = 4 Then
    If C1.Caption = x And C8.Caption = x And C6.Caption = o And C7.Caption = o Then b = 3
    If C3.Caption = x And C2.Caption = x And C8.Caption = o And C5.Caption = o Then b = 7
    If C7.Caption = x And C0.Caption = x And C2.Caption = o And C1.Caption = o Then b = 5
    If C5.Caption = x And C6.Caption = x And C0.Caption = o And C3.Caption = o Then b = 1
    If C1.Caption = x And C6.Caption = x And C8.Caption = o And C7.Caption = o Then b = 5
    If C3.Caption = x And C8.Caption = x And C2.Caption = o And C5.Caption = o Then b = 1
    If C7.Caption = x And C2.Caption = x And C0.Caption = o And C1.Caption = o Then b = 3
    If C5.Caption = x And C0.Caption = x And C6.Caption = o And C3.Caption = o Then b = 7
  End If
End If
End Sub

Public Sub h5()
If h > 3 Then
  If c = 6 Then
    If (C3.Caption = x And C6.Caption = x And C8.Caption = x) _
      And (C0.Caption = o And C5.Caption = o And C7.Caption = o) Then b = 1
    If (C0.Caption = x And C2.Caption = x And C3.Caption = x) _
      And (C1.Caption = o And C5.Caption = o And C6.Caption = o) Then b = 7
    If (C5.Caption = x And C6.Caption = x And C8.Caption = x) _
      And (C2.Caption = o And C3.Caption = o And C7.Caption = o) Then b = 1
    If (C0.Caption = x And C2.Caption = x And C5.Caption = x) _
      And (C1.Caption = o And C3.Caption = o And C8.Caption = o) Then b = 7
    If (C0.Caption = x And C6.Caption = x And C7.Caption = x) _
      And (C1.Caption = o And C3.Caption = o And C8.Caption = o) Then b = 5
    If (C2.Caption = x And C7.Caption = x And C8.Caption = x) _
      And (C1.Caption = o And C5.Caption = o And C6.Caption = o) Then b = 3
    If (C0.Caption = x And C1.Caption = x And C6.Caption = x) _
      And (C2.Caption = o And C3.Caption = o And C7.Caption = o) Then b = 5
    If (C1.Caption = x And C2.Caption = x And C8.Caption = x) _
      And (C0.Caption = o And C5.Caption = o And C7.Caption = o) Then b = 3
  End If
End If
End Sub
