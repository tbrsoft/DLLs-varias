VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0009FFDA&
   BorderStyle     =   0  'None
   Caption         =   "clsDiscosII Debug"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   593
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   674
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   330
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   1605
      Width           =   9600
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   330
      Picture         =   "form1.frx":0000
      Top             =   300
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Arrastra las fotos con el mouse (ESC para salir)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009EADAC&
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MouseDownX As Long
Dim MouseUpX As Long
Dim MouseMX As Long
Dim XDiferencia

Dim SuperX As Long

Dim ActivateCuliado As Boolean

Dim Discos2 As New clsDiscosII

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then End
End Sub

Private Sub Form_Load()
    Me.BackColor = vbBlack
    P.Left = ((Screen.Width / 15) / 2) - (P.Width / 2)
    P.Top = ((Screen.Height / 15) / 2) - (P.Height / 2)
    
    Dim Archivo As String
    Archivo = App.Path + "\img2\largo2.jpg"
    
    Discos2.IniciarGrafios P.hdc, 0, 0, 640, 480, Me
    Discos2.SetImagen Archivo
    
    PintarSuperX
End Sub

Private Sub P_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActivateCuliado = True
    MouseDownX = X / 15
End Sub


Private Sub P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ActivateCuliado = True Then
        MouseMX = X / 15
        If MouseDownX > MouseMX Then
            SuperX = SuperX + ((MouseDownX - MouseMX) * 0.05)
            XDiferencia = ((MouseDownX - MouseMX) * 0.05)
        Else
            SuperX = SuperX - ((MouseMX - MouseDownX) * 0.05)
            XDiferencia = ((MouseMX - MouseDownX) * 0.05)
        End If
        PintarSuperX
    End If
End Sub

Private Sub P_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    ActivateCuliado = False
    MouseUpX = X / 15
    
    'Inersia... si... ponele
    For i = (XDiferencia * 4) To 1 Step -1
        If MouseDownX > MouseMX Then
            SuperX = SuperX + (0.1 * i)
        Else
            SuperX = SuperX - (0.1 * i)
        End If
        PintarSuperX
    Next i
End Sub

'========================================================
Private Sub PintarSuperX()
    Discos2.PintarFrom SuperX, 0
    Discos2.Renderizar
    P.Refresh
End Sub
