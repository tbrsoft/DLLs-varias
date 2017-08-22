VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "PNG2"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6720
      Left            =   15
      Picture         =   "form.frx":0000
      ScaleHeight     =   6690
      ScaleWidth      =   8850
      TabIndex        =   0
      Top             =   15
      Width           =   8880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PNGView As New tbrPNG3

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    End
End Sub

Private Sub Form_Load()
    Dim Token As Long
    Token = PNGView.InitGDIPlus()
    PNGView.LoadPictureGDIPlus App.Path + "\boton.png", 120, 120, P.hDC
    PNGView.FreeGDIPlus Token
    'PNGView.FreeGDIPlus
    
End Sub

