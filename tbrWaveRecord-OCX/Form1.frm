VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00B4E0CA&
   Caption         =   "tbrWaveRecord"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin tbrWR_OCX.tbrWR tbrWR1 
      Left            =   4140
      Top             =   1860
      _ExtentX        =   847
      _ExtentY        =   847
      xDisp           =   0
      xLine           =   2
      xArchiv         =   "C:\tbrWaveRecord.wav"
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   2850
      Top             =   1005
   End
   Begin VB.PictureBox PV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   240
      Left            =   1980
      ScaleHeight     =   180
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   1725
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00B4E0CA&
      Caption         =   "<- Modo"
      Height          =   315
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1380
      Width           =   840
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
      Top             =   60
      Width           =   2310
   End
   Begin VB.ListBox Ld 
      Height          =   1035
      Left            =   30
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

Private Sub Command1_Click()
    Command2_Click
    
    'tbrWR1.Grabar 1, 1, "8000", 8
    tbrWR1.Grabar
    
    'E.Comenzar
    'E2.Comenzar
End Sub

Private Sub Command2_Click()
    tbrWR1.Detener
    'E.Detener
    'E2.Detener
End Sub

Private Sub Command3_Click()
    If E.Modo = 0 Then
        E.Modo = 1
    Else
        E.Modo = 0
    End If
End Sub

Private Sub Form_Load()
    'E.Dispositivo = 0
    'E.MIXERLINE = 2
    tbrWR1.Dispositivo = 0
    tbrWR1.Linea = 2
    
    tbrWR1.CargoDispositivos Ld
End Sub

Private Sub Ld_Click()
    'E.Dispositivo = Ld.ListIndex
    'E2.Dispositivo = Ld.ListIndex
    
    tbrWR1.Dispositivo = Ld.ListIndex
    tbrWR1.CargoLineas Ll
End Sub

Private Sub Ll_Click()
    tbrWR1.Linea = Ll.ListIndex
    'E.MIXERLINE = Ll.ListIndex
    'E.MIXERLINE = Ll.ListIndex
End Sub

Private Sub Timer1_Timer()
    PV.Cls
    Dim N As Single
    N = tbrWR1.VumeterNum
    'n=
    PV.Line (0, 0)-(N * PV.Width, PV.Height), vbYellow, BF
    
    If N > qMax Then
        qMax = N
    End If
    
    PV.Line ((qMax * PV.Width) - 15, 0)-((qMax * PV.Width) + 15, PV.Height), vbRed, BF
    
    If Timer > NumTim Then
        NumTim = Timer + 2
        qMax = 0
    End If
    
    'PV.CurrentX = 0
    'PV.CurrentY = 0
    'PV.Print N
End Sub
