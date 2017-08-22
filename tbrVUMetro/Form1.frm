VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6150
      Left            =   90
      ScaleHeight     =   6150
      ScaleWidth      =   1290
      TabIndex        =   9
      Top             =   630
      Width           =   1290
   End
   Begin VB.ListBox lstCantPics 
      Height          =   2205
      ItemData        =   "Form1.frx":0000
      Left            =   3600
      List            =   "Form1.frx":001C
      TabIndex        =   8
      Top             =   3660
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "MODO"
      Height          =   2535
      Left            =   3570
      TabIndex        =   2
      Top             =   750
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "4 AIWA"
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3 3PM"
         Height          =   285
         Index           =   3
         Left            =   270
         TabIndex        =   6
         Top             =   1350
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2 ESTEREO"
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   5
         Top             =   990
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1 ABAJO"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   660
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "0 ARRIBA"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   2160
      ScaleHeight     =   3090
      ScaleWidth      =   1230
      TabIndex        =   1
      Top             =   3660
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   465
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim V As New tbrSoftVumetro.tbrDrawVUM

Private Sub Form_Load()
    
    V.CantPic = 512
    V.CantCuadros = 20
    
    V.DefinePictureBox Picture1
    V.DefinePictureBox2 Picture2
    
    V.ModoVumetro = TresColoresEstereo
    
    V.Empezar
    
    Form2.Show
End Sub


Private Sub Command1_Click()
    V.Terminar
    Set V = Nothing
    Unload Me
End Sub


Private Sub lstCantPics_Click()
    V.CantPic = lstCantPics
    Dim T As Long
    T = Picture1.Width
    
    Picture1.Width = Picture1.Height
    Picture1.Height = T
    
    T = Picture2.Width
    Picture2.Width = Picture2.Height
    Picture2.Height = T
End Sub

Private Sub Option1_Click(Index As Integer)
    V.ModoVumetro = Index
    
    If Index = 3 Then V.ColorBase = vbBlack
        
    If Index = 0 Then
        V.DefinePictureBox Form2.Picture1
        V.DefinePictureBox2 Form2.Picture2
    Else
        V.DefinePictureBox Picture1
        V.DefinePictureBox2 Picture2
    End If
    
    V.NotifyResizeVUM
End Sub

Private Sub Picture1_Resize()
    V.NotifyResizeVUM
End Sub

Private Sub Picture2_Resize()
    V.NotifyResizeVUM
End Sub
