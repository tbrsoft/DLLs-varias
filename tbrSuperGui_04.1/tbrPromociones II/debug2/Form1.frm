VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   1740
      Top             =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TMvl As New tbrTextoMovil
Private Sub Form_Load()
    'Me.PaintPicture LoadPicture("c:\Fondo.bmp"), 0, 0
    TMvl.IniciarGrafios Me.hdc, 560, 20, 230, 30, 1
    Me.Cls
    TMvl.IniciarFuente Me, "Arial", 16, False, False, False, False, RGB(150, 150, 250)
    TMvl.SetTexto "Idea fea de texto que se mueve"
End Sub

Private Sub Timer1_Timer()
    TMvl.DibujarTexto
    Me.Refresh
End Sub
