VERSION 5.00
Begin VB.UserControl tbrPROGRESO 
   BackColor       =   &H00404000&
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   ScaleHeight     =   945
   ScaleWidth      =   5205
   Begin VB.Label lblPORC 
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin VB.Shape shBAR 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   105
      Top             =   540
      Width           =   255
   End
   Begin VB.Label lblTITULO 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo del proceso"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   5025
   End
End
Attribute VB_Name = "tbrPROGRESO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub ShowProceso(Titulo As String, PORC As Long)
    'mostrar una barra de progreso del frmPorceso
    lblTITULO = Titulo
    'por las dudas!!!
    If PORC > 100 Then PORC = 99
    shBAR.Width = PORC * lblTITULO.Width / 100
    lblPORC = CStr(PORC) + " %"
    
    lblTITULO.Refresh
    shBAR.Refresh
    lblPORC.Refresh
End Sub

