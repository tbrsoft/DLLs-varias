VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tbrTextoMovil"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   12300
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      Height          =   9000
      Left            =   75
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   1
      Top             =   390
      Width           =   12000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4575
      Top             =   75
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dale"
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim C As Long
Dim TMov As New tbrPromociones2
Private Sub Command1_Click()
    'Me.PaintPicture LoadPicture("C:\Fondo.bmp"), 0, 0
End Sub

Private Sub Command2_Click()
    'Me.Refresh
End Sub

Private Sub Form_Load()
    'P.PaintPicture LoadPicture("C:\Fondo.bmp"), 0, 0
    TMov.IniciarGrafios P.hdc, 25, 525, 250, 100, 1, 100
    TMov.IniciarFuente Me, "Arial", 8, False, False, False, False, RGB(250, 250, 250)
    
    TMov.AgregarPromo "Coman lombrices" + vbCrLf + "POSTA!" + vbCrLf + "Hacen bien"
    TMov.AgregarPromo "Los koalas son comunistas" + vbCrLf + "odian a los canguros" + vbCrLf + "por diversos mostivos"
    TMov.AgregarPromo "Los topos republicanos" + vbCrLf + "tiran misiles" + vbCrLf + "pero son ciegos"
    
    P.Refresh
    
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    'C = C - 1
    TMov.DibujarTexto
    P.Refresh
End Sub
