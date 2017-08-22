VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tbrEspectro"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Ch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cheto"
      Height          =   315
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   945
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Amplitud"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1650
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Espectro"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1380
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.ListBox lLine 
      Height          =   1035
      Left            =   2400
      TabIndex        =   4
      Top             =   300
      Width           =   2775
   End
   Begin VB.ListBox lDisp 
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   300
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detener"
      Height          =   375
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1410
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comenzar"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1410
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   4200
      Top             =   150
   End
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3600
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Linea:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   90
      Width           =   525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dispositivo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   60
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CGMA As Single
Private WithEvents clsRecorder  As WaveInRecorder
Attribute clsRecorder.VB_VarHelpID = -1
Private intSamples()            As Integer
Private clsVis                  As clsDraw

Private Sub clsRecorder_GotData(intBuffer() As Integer, lngLen As Long)
    intSamples = intBuffer
End Sub

Private Sub Command1_Click()
        
        If Not clsRecorder.StartRecord("44100", 2) Then
            MsgBox "No se puede comunicar con dispositivo", vbExclamation, "tbrEspectro"
        End If
        Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    If Not clsRecorder.StopRecord Then
        MsgBox "No se puede detener", vbExclamation
    End If
    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    'CGMA = -250
    Set clsRecorder = New WaveInRecorder
    Set clsVis = New clsDraw
    CargoDispositivos
    ReDim intSamples(FFT_SAMPLES - 1) As Integer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Command2_Click
    End
End Sub

Private Sub lDisp_Click()
    If Not clsRecorder.SelectDevice(lDisp.ListIndex) Then
        MsgBox "Ocurrio un error con este dispositivo!", vbExclamation, "Telefono"
        Exit Sub
    End If
    CargoLineas
End Sub

Private Sub Option1_Click()
    P.AutoRedraw = True
    P.Cls
End Sub

Private Sub Option2_Click()
    P.AutoRedraw = False
    P.Cls
End Sub

Private Sub Timer1_Timer()
    If Option1.value = True Then
        clsVis.DrawFrequencies intSamples, P, vbBlue
    Else
        clsVis.DrawAmplitudes intSamples, P
    End If
    'CGMA = CGMA + 10
    'If CGMA > 40 Then
    '    CGMA = -40
    'End If
    
    'fxBlur P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15)
    'fxEngrave P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15), 5
    'fxMosaic P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15), 2
    'fxGridelines P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15), vbBlack, 200, 1
    'fxRelief P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15)
    
    '_
    If Ch.value = 1 Then
        fxBlur P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15)
        fxLight P.hdc, 10, 10, vbWhite, 100, 50, 150
        fxScanlines P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15), vbBlack, 200, 1, True, False
    End If
    '¯
    
    'fxScanlines
    'P2.PaintPicture P.Image, 0, 0, P2.Width / 15, P2.Height / 15
End Sub

Sub CargoDispositivos()
    lDisp.Clear
    For i = 0 To clsRecorder.DeviceCount - 1
        lDisp.AddItem clsRecorder.DeviceName(i)
    Next
End Sub
Sub CargoLineas()
    lLine.Clear
    For i = 0 To clsRecorder.MixerLineCount - 1
        lLine.AddItem clsRecorder.MixerLineName(i)
    Next
End Sub

