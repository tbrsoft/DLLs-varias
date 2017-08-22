VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox P2 
      AutoSize        =   -1  'True
      Height          =   5685
      Left            =   4380
      ScaleHeight     =   5625
      ScaleWidth      =   4215
      TabIndex        =   2
      Top             =   660
      Width           =   4275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   8400
      TabIndex        =   1
      Top             =   30
      Width           =   1785
   End
   Begin VB.PictureBox P1 
      AutoSize        =   -1  'True
      Height          =   5685
      Left            =   30
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5625
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   660
      Width           =   4275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SP As New clsTbrScupePic

Private Sub Command1_Click()
    'SP.Levantar3D 0, 0, P1.Width, P1.Height, 0
End Sub

Private Sub Form_Load()
    SP.SetPP P1
End Sub

Private Sub P1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SP.Levantar3D X, Y, X + 400, Y + 3200, X + 1000, Y + 3200, 0
    'SP.Levantar3D X + 2000, Y, X + 2400, Y + 5200
    'SP.Levantar3D X, Y - 400, X + 2400, Y
    'SP.Levantar3D X, Y, X + 400, Y + 5200
    
    'P1.PaintPicture P1.Picture, X, Y, 300, 4000, X, Y, 1000, 4000
    'P1.PaintPicture P1.Picture, X, Y, 3000, 300, X, Y, 3000, 2000
    
End Sub

Private Sub P2_Click()
    P2.Cls
End Sub
