VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Abrir"
      Height          =   675
      Left            =   1110
      TabIndex        =   3
      Top             =   2940
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Grabar"
      Height          =   675
      Left            =   1110
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mostrar"
      Height          =   675
      Left            =   1110
      TabIndex        =   1
      Top             =   1380
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AgregarCFG"
      Height          =   675
      Left            =   1110
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TK As New tbrSelectKey.cl_tbrSoftSelectKey

Private Sub Command1_Click()
    Dim NN(2) As String
    
    NN(0) = InputBox("Inserte nombre de la config", , "No deje en blanco")
    NN(1) = InputBox("Inserte Descripcion", , "No deje en blanco")
    NN(2) = InputBox("Inserte Valor predeterminado (NUMERO!)", , "No deje en blanco")
    
    If Not IsNumeric(NN(2)) Then
        MsgBox "El ultimo debe ser numero idiota"
        Exit Sub
    End If
    
    TK.ADDcfg NN(0), CLng(NN(2)), NN(1)
    
    
End Sub

Private Sub Command2_Click()
    TK.ShowCFG TK
End Sub

Private Sub Command3_Click()
    TK.SaveCfg "c:\cfg.txt"
End Sub

Private Sub Command4_Click()
    TK.LoadCfg "c:\cfg.txt"
    MsgBox "Se han cargado " + CStr(TK.GetMaxCfg) + " configs"
End Sub
