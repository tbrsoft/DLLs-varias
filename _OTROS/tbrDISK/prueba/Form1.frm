VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   420
      Top             =   3000
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000C000&
      Height          =   705
      Left            =   150
      ScaleHeight     =   645
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   1020
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000080FF&
      Height          =   705
      Left            =   150
      ScaleHeight     =   645
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DR As New cslPitar

Private Sub Form_Load()
'    DIBUJAR
End Sub

Private Sub DIBUJAR()
    DR.DibijarElipseDegradee Me, 250, Me.Width / 30, Me.Height / 30, _
        Picture1.BackColor, Picture2.BackColor, 2
End Sub

'Private Sub Form_Load()
'
'
'    Dim S() As String
'    ReDim Preserve S(1): S(1) = "Un ejemplo de cancion"
'    ReDim Preserve S(2): S(2) = "Un ejemplo de cancion mas largo"
'    ReDim Preserve S(3): S(3) = "Un ejemplo de cancion maaaassss laaaargo"
'    ReDim Preserve S(4): S(4) = "Un ejemplo de cancion"
'    ReDim Preserve S(5): S(5) = "Un ejemplo de cancion"
'    ReDim Preserve S(6): S(6) = "Un ejemplo de cancion"
'    ReDim Preserve S(7): S(7) = "Un ejemplo de cancion"
'    ReDim Preserve S(8): S(8) = "Un ejemplo de cancion"
'    ReDim Preserve S(9): S(9) = "Un ejemplo de cancion"
''    ReDim Preserve S(10): S(10) = "Un ejemplo de cancion"
''    ReDim Preserve S(11): S(11) = "Un ejemplo de cancion"
''    ReDim Preserve S(12): S(12) = "Un ejemplo de cancion"
''    ReDim Preserve S(13): S(13) = "Un ejemplo de cancion"
''    ReDim Preserve S(14): S(14) = "Un ejemplo de cancion"
''    ReDim Preserve S(15): S(15) = "Un ejemplo de cancion"
''    ReDim Preserve S(16): S(16) = "Un ejemplo de cancion"
'
'    'DR.DibujarEllipsePattern Me, App.Path + "\disco.jpg", 100, S
'
'    DR.DibijarElipseDegradee Me, 50, 300, 300, S
'End Sub

Private Sub Picture1_Click()
    Dim CD As New CommonDialog
    CD.RGBResult = Picture1.BackColor
    CD.ShowColor
    
    Picture1.BackColor = CD.RGBResult
    
    DIBUJAR
End Sub

Private Sub Picture2_Click()
    Dim CD As New CommonDialog
    CD.RGBResult = Picture2.BackColor
    CD.ShowColor
    
    Picture2.BackColor = CD.RGBResult
    
    DIBUJAR
End Sub

