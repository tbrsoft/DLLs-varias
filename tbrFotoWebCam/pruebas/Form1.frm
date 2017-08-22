VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2700
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   960
      Width           =   825
   End
   Begin VB.CommandButton Command9 
      Caption         =   "eventos ON"
      Height          =   555
      Left            =   3960
      TabIndex        =   20
      Top             =   2910
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Capacidades"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   270
      TabIndex        =   13
      Top             =   2010
      Width           =   3585
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "HasOverlay"
         Enabled         =   0   'False
         ForeColor       =   &H8000000A&
         Height          =   345
         Left            =   60
         TabIndex        =   19
         Top             =   870
         Width           =   1965
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "DriverSuppliesPalettes"
         Enabled         =   0   'False
         ForeColor       =   &H8000000A&
         Height          =   345
         Left            =   60
         TabIndex        =   18
         Top             =   570
         Width           =   1965
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Captura inicializada ok"
         Enabled         =   0   'False
         ForeColor       =   &H8000000A&
         Height          =   345
         Left            =   60
         TabIndex        =   17
         Top             =   270
         Width           =   1905
      End
      Begin VB.CommandButton cmdDlgVideoSource 
         Caption         =   "DlgVideoSource"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2010
         TabIndex        =   16
         Top             =   900
         Width           =   1500
      End
      Begin VB.CommandButton cmdDlgVideoFormat 
         Caption         =   "DlgVideoFormat"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2010
         TabIndex        =   15
         Top             =   600
         Width           =   1500
      End
      Begin VB.CommandButton cmdDlgVideoDisplay 
         Caption         =   "DlgVideoDisplay"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2010
         TabIndex        =   14
         Top             =   300
         Width           =   1500
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Obtener Capacidades"
      Height          =   585
      Left            =   300
      TabIndex        =   10
      Top             =   1410
      Width           =   1500
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Obtener Descripcion"
      Height          =   585
      Left            =   360
      TabIndex        =   9
      Top             =   60
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   1530
   End
   Begin VB.CommandButton Command6 
      Caption         =   "fotoSSSS"
      Height          =   585
      Left            =   7410
      TabIndex        =   8
      Top             =   2610
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   1845
      Left            =   3960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "Form1.frx":030A
      Top             =   690
      Width           =   5835
   End
   Begin VB.CommandButton Command4 
      Caption         =   "No mostrar (no desconecta)"
      Height          =   555
      Left            =   5550
      TabIndex        =   6
      Top             =   3660
      Width           =   1515
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mostrar"
      Height          =   555
      Left            =   3990
      TabIndex        =   5
      Top             =   3660
      Width           =   1515
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Conectar driver"
      Height          =   585
      Left            =   330
      TabIndex        =   4
      Top             =   690
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "foto"
      Height          =   585
      Left            =   5760
      TabIndex        =   3
      Top             =   2610
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   4980
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   4260
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "salir"
      Height          =   5775
      Left            =   9900
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   120
      ScaleHeight     =   3600
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   4260
      Width           =   4800
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   525
      Left            =   5820
      TabIndex        =   12
      Top             =   90
      Width           =   3945
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1860
      TabIndex        =   11
      Top             =   120
      Width           =   3885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents TWC As tbrCameraDrivers.tbrWEBCAM
Attribute TWC.VB_VarHelpID = -1

Private Sub cmdDlgVideoDisplay_Click()
    TWC.Mostrar_DlgVideoDisplay
End Sub

Private Sub cmdDlgVideoFormat_Click()
    TWC.Mostrar_DlgVideoFormat_Click
End Sub

Private Sub cmdDlgVideoSource_Click()
    TWC.Mostrar_DlgVideoSource_Click
End Sub

Private Sub Command1_Click()
    TWC.IniciarEvento False, evStatus
    TWC.IniciarEvento False, evError
    DoEvents
    TWC.Preview_Detener
    TWC.Driver_Desconectar
    Unload Me
End Sub

Private Sub Command2_Click()
    Text2.Text = Text2.Text + "pide foto " + CStr(Timer) + vbCrLf
    TWC.SacarFoto_ClipBoard
    Text2.Text = Text2.Text + "llega foto " + CStr(Timer) + vbCrLf
    'Picture2.PaintPicture Clipboard.GetData, 0, 0, Picture1.Width, Picture1.Height
    Picture2.Picture = Clipboard.GetData
    Text2.Text = Text2.Text + "pegue foto " + CStr(Timer) + vbCrLf
    Clipboard.Clear
    Text2.Text = Text2.Text + "clip clean " + CStr(Timer) + vbCrLf
End Sub

Private Sub Command3_Click()
    On Local Error GoTo myErr

    TWC.Preview_Iniciar

myErr:
    TWC_Status Err.Description
End Sub

Private Sub Command4_Click()
    TWC.Preview_Detener
End Sub

Private Sub Command5_Click()
    On Local Error GoTo myErr
    DoEvents
    
    If TWC.Driver_Conectar(Picture1.hWnd, 320, 240) <> 0 Then
        MsgBox "No se puedo conectar el driver!"
    End If
    
    Exit Sub
    
myErr:
    TWC_Status Err.Description
End Sub

Private Sub Command6_Click()
    If Timer1.Interval = 0 Then
        Timer1.Interval = 100: Command6.Caption = "Detener"
    Else
        Timer1.Interval = 0: Command6.Caption = "Iniciar"
    End If
End Sub

Private Sub Command7_Click()
    'aca lee el driver solamente
    TWC.GetDriverDescription
    Label1.Caption = "Driver Name: " + TWC.GetDriverName
    Label2.Caption = "Driver Version: " + TWC.GetDriverVersion
End Sub

Private Sub Command8_Click()
    
    TWC.GetCapabilities
    
    cmdDlgVideoDisplay.Enabled = TWC.Puede_DlgVideoDisplay
    cmdDlgVideoFormat.Enabled = TWC.Puede_DlgVideoFormat
    cmdDlgVideoSource.Enabled = TWC.Puede_DlgVideoSource

    If TWC.InicioOK Then Check1.Value = 1
    
    If TWC.Puede_SoportarPaletas Then Check2.Value = 1
    If TWC.Puede_Overlay Then Check1.Value = 1
End Sub

Private Sub Command9_Click()
    TWC.mHwndMsgSET Text1.hWnd
    TWC.IniciarEvento True, evError
    TWC.IniciarEvento True, evStatus
End Sub

Private Sub Form_Load()
    Set TWC = New tbrCameraDrivers.tbrWEBCAM
    Picture2.AutoRedraw = True
End Sub

Private Sub Text1_Change()
    Text2 = Text2 + Text1 + vbCrLf
    Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
    TWC.SacarFoto_ClipBoard
    'Picture2.PaintPicture Clipboard.GetData, 0, 0, Picture1.Width, Picture1.Height
    Picture2.Picture = Clipboard.GetData
    Clipboard.Clear
End Sub

Private Sub TWC_Error(DetalleError As String, IdError As Long)
    Text2.Text = Text2.Text + "(error " + CStr(IdError) + ") " + DetalleError + vbCrLf
End Sub

Private Sub TWC_Status(DetalleStatus As String)
    Text2.Text = Text2.Text + "ST: " + DetalleStatus + vbCrLf
End Sub
