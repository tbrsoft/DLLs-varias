VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTMP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0017B4FF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   1050
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   1
      Top             =   915
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   915
   End
   Begin VB.Label lbl 
      BackColor       =   &H009EADAC&
      Caption         =   "imagenes."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1050
      TabIndex        =   2
      Top             =   90
      Width           =   7020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Reflejo As clsGeneradoReflex
Attribute Reflejo.VB_VarHelpID = -1
Dim lasImagenes() As String

Private Sub Command1_Click()
    Set Reflejo = New clsGeneradoReflex
    CargarArchivos App.Path + "\img\", "*.jpg", lasImagenes()
    
    picTMP.Width = 10
    picTMP.Height = 10
    
    Reflejo.GenerarImagen lasImagenes(), 280, 200, picTMP.hdc
    Reflejo.GrabarEnArchivo App.Path + "\chance.bmp", picTMP
    picTMP.Visible = True
    Reflejo.CerrarGraficos

End Sub

Sub CargarArchivos(qPath As String, qArchivosBuscar As String, qMatriz() As String)
    Dim ArchivosEncontrados As String
    Dim ix As Long
    
    ReDim qMatriz(0)
    ArchivosEncontrados = Dir(qPath + qArchivosBuscar)
    While ArchivosEncontrados <> ""
        ix = ix + 1
        ReDim Preserve qMatriz(ix)
        qMatriz(ix) = qPath + ArchivosEncontrados
        ArchivosEncontrados = Dir
    Wend
End Sub

Private Sub Form_Load()
    Set qFormularioAuxiliar = Me
End Sub

Private Sub Reflejo_CargandoArchivo(qArchivo As String, qPorcentajeVoy As Long)
    lbl.Caption = CStr(qPorcentajeVoy) + "%"
    lbl.Caption = lbl.Caption + vbCrLf + "*Cargando: " + qArchivo
    lbl.Refresh
End Sub
