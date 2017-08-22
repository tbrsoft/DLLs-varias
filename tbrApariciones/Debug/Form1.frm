VERSION 5.00
Object = "*\A..\..\TBRAPA~1\tbrApariciones.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00D3DEDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00D3DEDE&
      Caption         =   "Marca de Agua"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   390
      Width           =   1635
   End
   Begin tbrAparicionesP.tbrApariciones tbrA 
      Left            =   4980
      Top             =   2460
      _ExtentX        =   953
      _ExtentY        =   953
      xTime           =   100
      xFrame          =   15
      xColoT          =   65280
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D3DEDE&
      Caption         =   "Desaparecer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1830
      Width           =   1425
   End
   Begin VB.TextBox FR 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7155
      TabIndex        =   8
      Text            =   "15"
      Top             =   1290
      Width           =   735
   End
   Begin VB.TextBox TI 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7155
      TabIndex        =   6
      Text            =   "200"
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D3DEDE&
      Caption         =   "Aparecer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5355
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1830
      Width           =   1095
   End
   Begin VB.PictureBox P2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   5340
      Picture         =   "Form1.frx":2E7A
      ScaleHeight     =   2190
      ScaleWidth      =   4680
      TabIndex        =   4
      Top             =   2340
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3225
      Left            =   690
      Picture         =   "Form1.frx":2448C
      ScaleHeight     =   3225
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   1125
      Width           =   3750
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Height          =   150
      Left            =   7365
      TabIndex        =   12
      Top             =   1635
      Width           =   150
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Transparente:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5355
      TabIndex        =   11
      Top             =   1575
      Width           =   1950
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frame Rate:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5370
      TabIndex        =   9
      Top             =   1290
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time Interval:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5370
      TabIndex        =   7
      Top             =   990
      Width           =   1395
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1410
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   870
      Width           =   2715
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":271E4
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5295
      TabIndex        =   3
      Top             =   4515
      Width           =   4965
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-PictureDestino: Este es el Picture a donde se imprime la imagen, debe tener una imagen en la propiedad Picture"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   465
      TabIndex        =   2
      Top             =   4425
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "El OCX es intermediario entre 2 PictureBox [QUE DEBEN TENER AUTOREDRAW = TRUE]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   8310
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   4275
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   4785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Si no se aclara Width o Height entonces se usa el alto y ancho de la ImagenFuente

Private Sub Command1_Click()
    tbrA.Aparecer P1, P2, 0, 0, True
End Sub

Private Sub Command2_Click()
    tbrA.Aparecer P1, P2, 0, 0, False
End Sub

Private Sub Command3_Click()
    Dim N As Long
    N = InputBox("Nivel de Marca de Agua. Numero entre 0 y 255", "Marca de Agua", 130)
    P1.Cls
    tbrA.MarcaDeAgua P1, vbWhite, N
End Sub

Private Sub FR_Change()
    tbrA.FrameRate = Val(FR.Text)
End Sub

Private Sub TI_Change()
    tbrA.TimeInterval = Val(TI.Text)
End Sub
    

