VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tbrMP3Enc"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Paux 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   165
      ScaleHeight     =   285
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   690
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "archivo"
      Height          =   285
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   270
      Width           =   1095
   End
   Begin VB.PictureBox pBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   75
      Picture         =   "form1.frx":058A
      ScaleHeight     =   450
      ScaleWidth      =   3000
      TabIndex        =   3
      Top             =   585
      Width           =   3030
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "to mp3"
      Height          =   480
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   630
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   3030
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo a Encriptar"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   30
      Width           =   1605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tbrMP3var As tbrMP3Enc.tbrMP3EncDll
Attribute tbrMP3var.VB_VarHelpID = -1
Dim WithEvents GetEvento As tbrMP3Enc.clsLlamarEvento
Attribute GetEvento.VB_VarHelpID = -1

Dim CD As New CommonDialog

Sub SetPorciento(Pje As Integer)
    Dim Ancho As Long
    Dim Alto As Long
    Ancho = (Pje / 100) * pBarra.Width
    Alto = pBarra.Height
    If Ancho < 1 Then Exit Sub
    pBarra.PaintPicture Paux.Image, 0, 0, Ancho, Alto, , , , , vbSrcAnd
    pBarra.Refresh
End Sub

Private Sub Command1_Click()
    Dim Res As Integer
    tbrMP3var.Iniciar f_Kbps_128_Default
    If Text1.Text <> "" Then
        Res = tbrMP3var.Encode(Text1.Text)
        If Res <> -1 Then
            MsgBox "Codificado Correctamente", vbInformation, "Proceso terminado"
        End If
    Else
        MsgBox "Elija un archivo!", vbCritical, "tbrError"
    End If
End Sub

Private Sub Command2_Click()
    Dim Carp As String
    Carp = GetSetting("tbrMP3Encr", "Memory", "Archivo", "C:\")
    CD.InitDir = Carp
    
    
    CD.FileName = ""
    CD.Filter = "Archivos WAV|*.wav"
    CD.DialogTitle = "Abrir"
    CD.ShowOpen
    
    If CD.FileName <> "" Then
        If Dir(CD.FileName) <> "" Then
            Text1.Text = CD.FileName
            Carp = Mid(CD.FileName, 1, InStrRev(CD.FileName, "\"))
            SaveSetting "tbrMP3Encr", "Memory", "Archivo", Carp
        End If
    End If
End Sub

Private Sub Form_Load()
    Set tbrMP3var = New tbrMP3Enc.tbrMP3EncDll
    Set GetEvento = tbrMP3var.GetEventos
End Sub

Private Sub GetEvento_Estado(Porcentaje As Integer)
    SetPorciento Porcentaje
End Sub
