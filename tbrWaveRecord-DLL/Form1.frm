VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00B4E0CA&
   Caption         =   "tbrWaveRecord"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PV 
      Height          =   2715
      Left            =   5220
      ScaleHeight     =   2655
      ScaleWidth      =   3405
      TabIndex        =   4
      Top             =   210
      Width           =   3465
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   2850
      Top             =   1005
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00B4E0CA&
      Caption         =   "Detener"
      Height          =   270
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00B4E0CA&
      Caption         =   "Grabar"
      Height          =   270
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1140
      Width           =   1050
   End
   Begin VB.ListBox Ll 
      Height          =   1035
      Left            =   2430
      TabIndex        =   1
      Top             =   45
      Width           =   2310
   End
   Begin VB.ListBox Ld 
      Height          =   1035
      Left            =   53
      TabIndex        =   0
      Top             =   45
      Width           =   2310
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qMax As Single
Dim NumTim As Single
Dim tbrWR1 As New tbrWR2

Private Sub Command1_Click()
    Command2_Click
    tbrWR1.Grabar
End Sub

Private Sub Command2_Click()
    tbrWR1.Detener
End Sub

Private Sub Form_Load()
    tbrWR1.Dispositivo = 0
    tbrWR1.Linea = 2
    tbrWR1.CargoDispositivos Ld
End Sub

Private Sub Ld_Click()
    tbrWR1.Dispositivo = Ld.ListIndex
    tbrWR1.CargoLineas Ll
End Sub

Private Sub Ll_Click()
    tbrWR1.Linea = Ll.ListIndex
End Sub
