VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "cTextEx Demo"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pTest 
      AutoRedraw      =   -1  'True
      Height          =   1815
      Left            =   585
      Picture         =   "fMain.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   915
      Width           =   7335
   End
   Begin VB.CommandButton cPBox 
      Caption         =   "DoIT"
      Height          =   435
      Left            =   630
      TabIndex        =   1
      Top             =   135
      Width           =   1215
   End
   Begin VB.CommandButton cExit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   7260
      TabIndex        =   0
      Top             =   3420
      Width           =   1095
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aFont As New pTextEx_Demo.clsMain


Private Sub cPBox_Click()
    aFont.Iniciar "Arial", 60, True, False, False, False, RGB(100, 100, 200)
    aFont.Dibujar "HOLA", pTest.hDC, 20, 20, 300, 100
    pTest.Refresh
End Sub
