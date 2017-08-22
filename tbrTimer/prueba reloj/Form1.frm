VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      ItemData        =   "Form1.frx":0000
      Left            =   3300
      List            =   "Form1.frx":003D
      TabIndex        =   1
      Top             =   90
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "empezar reloj"
      Height          =   465
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents T As tbrTimer.clsTimer
Attribute T.VB_VarHelpID = -1
Dim WithEvents T2 As tbrTimer.clsTimer
Attribute T2.VB_VarHelpID = -1
Dim WithEvents T3 As tbrTimer.clsTimer
Attribute T3.VB_VarHelpID = -1
Dim WithEvents T4 As tbrTimer.clsTimer
Attribute T4.VB_VarHelpID = -1
Dim WithEvents T5 As tbrTimer.clsTimer
Attribute T5.VB_VarHelpID = -1
Dim WithEvents T6 As tbrTimer.clsTimer
Attribute T6.VB_VarHelpID = -1
Dim WithEvents T7 As tbrTimer.clsTimer
Attribute T7.VB_VarHelpID = -1
Dim WithEvents T8 As tbrTimer.clsTimer
Attribute T8.VB_VarHelpID = -1
Dim WithEvents T9 As tbrTimer.clsTimer
Attribute T9.VB_VarHelpID = -1
Dim WithEvents T10 As tbrTimer.clsTimer
Attribute T10.VB_VarHelpID = -1

Dim H(10) As Long

Private Sub Command1_Click()
    T.Interval = 100: T.Enabled = True
    T2.Interval = 200: T2.Enabled = True
    T3.Interval = 300: T3.Enabled = True
    T4.Interval = 400: T4.Enabled = True
    T5.Interval = 500: T5.Enabled = True
    T6.Interval = 600: T6.Enabled = True
    T7.Interval = 700: T7.Enabled = True
    T8.Interval = 800: T8.Enabled = True
    T9.Interval = 900: T9.Enabled = True
    T10.Interval = 1000: T10.Enabled = True
End Sub

Private Sub Form_Load()
    Set T = New tbrTimer.clsTimer
    Set T2 = New tbrTimer.clsTimer
    Set T3 = New tbrTimer.clsTimer
    Set T4 = New tbrTimer.clsTimer
    Set T5 = New tbrTimer.clsTimer
    Set T6 = New tbrTimer.clsTimer
    Set T7 = New tbrTimer.clsTimer
    Set T8 = New tbrTimer.clsTimer
    Set T9 = New tbrTimer.clsTimer
    Set T10 = New tbrTimer.clsTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    T.Enabled = False
End Sub

Private Sub T_Timer(): H(0) = H(0) + 1: List1.List(0) = "0:" + CStr(H(0)): End Sub

Private Sub T2_Timer(): H(1) = H(1) + 1: List1.List(1) = "1:" + CStr(H(1)): End Sub

Private Sub T3_Timer(): H(2) = H(2) + 1: List1.List(2) = "2:" + CStr(H(2)): End Sub

Private Sub T4_Timer(): H(3) = H(3) + 1: List1.List(3) = "3:" + CStr(H(3)): End Sub

Private Sub T5_Timer(): H(4) = H(4) + 1: List1.List(4) = "4:" + CStr(H(4)): End Sub

Private Sub T6_Timer(): H(5) = H(5) + 1: List1.List(5) = "5:" + CStr(H(5)): End Sub

Private Sub T7_Timer(): H(6) = H(6) + 1: List1.List(6) = "6:" + CStr(H(6)): End Sub

Private Sub T8_Timer(): H(7) = H(7) + 1: List1.List(7) = "7:" + CStr(H(7)): End Sub

Private Sub T9_Timer(): H(8) = H(8) + 1: List1.List(8) = "8:" + CStr(H(8)): End Sub

Private Sub T10_Timer(): H(9) = H(9) + 1: List1.List(9) = "9:" + CStr(H(9)): End Sub
