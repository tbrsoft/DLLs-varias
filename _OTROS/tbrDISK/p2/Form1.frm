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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim DR As New tbrDiskDll.cslPitar
    Dim S() As String
    ReDim Preserve S(1): S(1) = "Un ejemplo de cancion"
    ReDim Preserve S(2): S(2) = "Un ejemplo de cancion mas largo"
    ReDim Preserve S(3): S(3) = "Un ejemplo de cancion maaaassss laaaargo"
    ReDim Preserve S(4): S(4) = "Un ejemplo de cancion"
    ReDim Preserve S(5): S(5) = "Un ejemplo de cancion"
    ReDim Preserve S(6): S(6) = "Un ejemplo de cancion"
    ReDim Preserve S(7): S(7) = "Un ejemplo de cancion"
    ReDim Preserve S(8): S(8) = "Un ejemplo de cancion"
    ReDim Preserve S(9): S(9) = "Un ejemplo de cancion"
    ReDim Preserve S(10): S(10) = "Un ejemplo de cancion"
    ReDim Preserve S(11): S(11) = "Un ejemplo de cancion"
    ReDim Preserve S(12): S(12) = "Un ejemplo de cancion"
    ReDim Preserve S(13): S(13) = "Un ejemplo de cancion"
    ReDim Preserve S(14): S(14) = "Un ejemplo de cancion"
    ReDim Preserve S(15): S(15) = "Un ejemplo de cancion"
    ReDim Preserve S(16): S(16) = "Un ejemplo de cancion"
    
    DR.DibujarEllipsePattern Me, App.Path + "\disco.jpg", Me.Width / 2, S
    
End Sub

