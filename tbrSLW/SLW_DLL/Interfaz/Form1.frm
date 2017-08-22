VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tbrSLW - Server"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Send Test"
      Height          =   495
      Left            =   9120
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox picCFG 
      Height          =   1335
      Left            =   5040
      ScaleHeight     =   1275
      ScaleWidth      =   5235
      TabIndex        =   8
      Top             =   4440
      Width           =   5295
      Begin VB.CommandButton Command3 
         Caption         =   "Guardar Configuracion"
         Height          =   495
         Left            =   3600
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox tPath 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "c:\"
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox tPuerto 
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Text            =   "8881"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Path de Licencias"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Puerto"
         Height          =   195
         Left            =   2040
         TabIndex        =   10
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5280
      Top             =   120
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   4800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apagar"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encender"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.Label lEst 
         AutoSize        =   -1  'True
         Caption         =   "en espera"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.TextBox tLog 
      Height          =   3495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   840
      Width           =   10215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Log"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conexiones multpiples simultaneas estan prohibidas"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   4485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents SLW As clsSLW
Attribute SLW.VB_VarHelpID = -1
Dim PathLic As String
Dim Puerto As Long


Private Sub Command1_Click()
    SLW.ComenzarServicio Puerto
End Sub

Private Sub Command2_Click()
    SLW.DetenerServicio
End Sub

Private Sub Command3_Click()
    PathLic = tPath
    Puerto = Val(tPuerto)
    
    SaveSetting "tbrSLW_gui", "cfg", "PathLicencia", PathLic
    SaveSetting "tbrSLW_gui", "cfg", "Puerto", tPuerto
End Sub

Private Sub Command4_Click()
    tLog.Text = tLog.Text + "Datos Enviados" + vbCrLf
    'WS.SendData "001//Maribel se durmio, vamos a cantarle por que se undio"
    WS.SendData "099//Maribel se durmio, vamos a cantarle por que se undio"
End Sub

Private Sub Form_Load()
    Set SLW = New clsSLW
    'Leo las configuraciones
    PathLic = GetSetting("tbrSLW_gui", "cfg", "PathLicencia", App.Path + "\Licencias\")
    Puerto = Val(GetSetting("tbrSLW_gui", "cfg", "Puerto", "8881"))
    tPath = PathLic
    tPuerto = CStr(Puerto)
        
    
    If SLW.InicializarPath(PathLic) = 1 Then
        'El directorio no existe NI SE PUEDE CREAR
        'MsgBox "El Path para las licencias no existe!" + vbCrLf + _
            "El programa no va a funcionar, cambie el Path por uno existente" _
            , vbInformation, "Path no existente"
            
        PathLic = App.Path + "\Licencias\"
        tPath = PathLic
        SLW.InicializarPath (PathLic)
    End If
    SLW.InicializarSocket WS
    
    'Iniciar Automaticamente
    Command1_Click
End Sub

Private Sub SLW_Suceso(Suceso As String)
    tLog.Text = tLog.Text + Suceso + vbCrLf
End Sub

Private Sub Timer1_Timer()
    lEst = SLW.GetStrEstado
End Sub

'========================================================
'(!) Importante Conectar estos 2 eventos!
'========================================================
Private Sub WS_ConnectionRequest(ByVal requestID As Long)
    SLW.ConnectionRequest requestID
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
    SLW.DataArrival bytesTotal
End Sub
