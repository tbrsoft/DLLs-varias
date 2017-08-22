VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   7500
      TabIndex        =   3
      Top             =   4080
      Width           =   1485
   End
   Begin VB.ListBox List2 
      Height          =   3450
      IntegralHeight  =   0   'False
      Left            =   5970
      TabIndex        =   2
      Top             =   540
      Width           =   4395
   End
   Begin VB.ListBox List1 
      Height          =   3420
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   1
      Top             =   570
      Width           =   5865
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Elegir carpeta"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FOL As String
Dim RES() As String
Dim FSO As New Scripting.FileSystemObject

Private Sub Command1_Click()
    Dim CM As New CommonDialog
    CM.InitDir = "C:\"
    CM.ShowFolder
    
    FOL = CM.InitDir
    
    If FOL = "" Or LCase(FOL = "c:\") Then Exit Sub
    
    Dim BR As New tbrPaths.clspaths
    
    BR.LeerTodo FOL, False, False, "*.mn1"
    
    RES = BR.GetLista
    
    List1.Clear: List2.Clear
    
    Dim J As Long, TMP As String
    For J = 1 To UBound(RES)
        List1.AddItem RES(J)
        TMP = FSO.GetBaseName(RES(J))
        If LCase(TMP) <> "tapa" Then
            If LCase(FSO.GetExtensionName(RES(J))) <> "mn1" Then List2.AddItem "*****"
            List2.AddItem TMP
        End If
    Next J
    
End Sub

Private Sub Command3_Click()
    Dim TE As TextStream, J As Long
    Set TE = FSO.CreateTextFile("c:\Lista_Karaokes.txt", True)
    
        For J = 1 To List2.ListCount - 1
            TE.WriteLine List2.List(J)
        Next J
    
    TE.Close
    
End Sub
