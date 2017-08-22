VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tbrTextoMovil"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7290
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0009FFDA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "AxeBlack"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3900
      Left            =   180
      ScaleHeight     =   3900
      ScaleWidth      =   6780
      TabIndex        =   2
      Top             =   900
      Width           =   6780
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3720
      Top             =   90
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Escribir"
      Height          =   375
      Left            =   1470
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dibujar Fondo"
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As Long
Dim TMov As New tbrTextoMovil
Private Sub Command1_Click()
    'Me.PaintPicture LoadPicture("C:\Fondo.bmp"), 0, 0
End Sub

Private Sub Command2_Click()
    'Me.Refresh
End Sub

Private Sub Form_Load()
    P.PaintPicture LoadPicture("C:\Fondo.bmp"), 0, 0
    TMov.Iniciar P, 130, 130, 150, 100
    TMov.ElTexto = "Hola Amigo"

End Sub

Private Sub Timer1_Timer()
    C = C - 1
    TMov.DibujarTexto C
    P.Refresh
End Sub
