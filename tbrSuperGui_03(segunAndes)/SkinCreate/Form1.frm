VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   4815
      Left            =   150
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   467
      TabIndex        =   0
      Top             =   150
      Width           =   7065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim G As New tbrSuperGui_3.clsGUI
    Dim frm As tbrSuperGui_3.ObjFullPadre
    
    Set frm = G.MNG.AddPadre("frmManu")
    
    Dim Obj As New tbrSuperGui_3.objFULL
    Obj.oSimple.X = 22
    Obj.oSimple.Y = 55
    Obj.oSimple.W = 240
    Obj.oSimple.H = 60
    
    Obj.Tipo = en_clsLabel
    Obj.CreateManu
    
    frm.AppendSGO Obj
    
    Obj.oSimple.SetProp "fuente", "verdana"
    
    frm.sHDC = P1.hdc
    
    frm.INIT_GRAPH
    
End Sub
