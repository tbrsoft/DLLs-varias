VERSION 5.00
Object = "{0371DBBE-C4D8-44B1-BFEE-712E91095894}#10.0#0"; "tbrListaConfig.ocx"
Begin VB.Form frmTest 
   Caption         =   "Pruebas"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   Begin tbrListaConfig_CTL.ctlFullCFG CCFG 
      Height          =   2025
      Left            =   1740
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3572
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cCFG_Fin(Rstrt As Long)
    
    If Rstrt = 2 Then REINICIAR_PC
    If Rstrt = 1 Then cCFG.PlayRockola
    
    Unload Me
    End '?
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyZ: cCFG.GoLeft
        Case vbKeyX: cCFG.GoRight
        Case vbKeyReturn: cCFG.GoOK
    End Select
End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True
    
End Sub

Private Sub Form_Resize()
    cCFG.Top = 190
    cCFG.Left = 60
    cCFG.Width = Me.Width - 350
    cCFG.Height = Me.Height - 500
End Sub

Public Sub DoClose() 'mentira, es abrir
    
    'SEGUIRAQUI
    'tuner segun software
    'Select Case cCFG.SoftNow
    '    Case ""
   '
   '
    'End Select
    Me.Show 1

End Sub
