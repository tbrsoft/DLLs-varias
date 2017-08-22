VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "ConfigTester"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   10620
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView TV 
      Height          =   750
      Left            =   7020
      TabIndex        =   1
      Top             =   5700
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   75
      ScaleHeight     =   559
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   696
      TabIndex        =   0
      Top             =   90
      Width           =   10470
   End
   Begin VB.Label Label1 
      BackColor       =   &H00BB00F9&
      Caption         =   "Label1"
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   8535
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conf As New clsConfig

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
            Conf.ComandoAdelante
        Case vbKeyLeft
            Conf.ComandoAtras
        Case 13
            Conf.ComandoEntrar
    End Select
    
    Conf.Renderizar
    P.Refresh
End Sub

Private Sub Form_Load()
    Conf.IniciarNodos TV, "R"
    
    Conf.Nodos.Add , , "R", "Raiz"
    Conf.Nodos.Add "R", tvwChild, "H1", "Hijo1*PNG\01.png"
    Conf.Nodos.Add "R", tvwChild, "H2", "Hijo2*PNG\02.png"
    Conf.Nodos.Add "R", tvwChild, "H3", "Hijo3"
    Conf.Nodos.Add "R", tvwChild, "H4", "Hijo4*PNG\04.png"
    Conf.Nodos.Add "R", tvwChild, "H5", "Hijo5"

    Conf.Nodos.Add "H2", tvwChild, "N1", "Nene1"
    Conf.Nodos.Add "H2", tvwChild, "N2", "Nene2"
    Conf.Nodos.Add "H2", tvwChild, "N3", "Nene3"
    
    Conf.Nodos.Add "N3", tvwChild, "Ni1", "Nieto1"
    Conf.Nodos.Add "N3", tvwChild, "Ni2", "Nieto2"
    Conf.Nodos.Add "N3", tvwChild, "Ni3", "Nieto3"
    
    Conf.IniciarFuente Me, "Arial", 12, False, False, False, False, vbWhite, RGB(50, 50, 50)
    Conf.IniciarGrafios P.hdc, 10, 10, App.path + "\conf.jpg", App.path + "\PNG\Selected.png", App.path + "\PNG\"
    Conf.CargarNodos
    
    Conf.Renderizar
    P.Refresh
End Sub
