VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   10095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar"
      Height          =   675
      Left            =   390
      TabIndex        =   3
      Top             =   5250
      Width           =   1665
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   270
      TabIndex        =   2
      Top             =   870
      Width           =   2025
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   675
      Left            =   330
      TabIndex        =   0
      Top             =   90
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fol1 As Folder 'cada uno de los sistemas disponibles
Dim Fol2 As Folder 'cada uno de los sistemas disponibles
Dim FSO As New Scripting.FileSystemObject
Dim AP As String
Dim THIS As String

Private Sub Command1_Click()
    Set Fol1 = FSO.GetFolder(AP)
    
    List2.Clear
    
    For Each Fol2 In Fol1.SubFolders
        List2.AddItem Fol2.Name + " (" + CStr(Fol2.Files.Count) + ")"
    Next
    
End Sub

Private Sub Command2_Click()
    Dim CE1 As New tbrCrypto.Crypt    'encripta el nombre del programa
    Dim progEncr As String
    progEncr = CE1.EncryptString(eMC_Blowfish, THIS, "Cerrar sistema", True)
    Text1.Text = ""
    Dim K As Long, renglon As String
    Dim codPlaca As String, codPlacaEncr As String
    For K = 0 To List1.ListCount - 1
        codPlaca = List1.List(K)
        codPlacaEncr = CE1.EncryptString(eMC_Blowfish, codPlaca, "Cerrar sistema", True)
        
        renglon = "'" + codPlaca + " AUTOM. " + THIS + vbCrLf + _
            "If cOd7 = dcr(" + Chr(34) + codPlacaEncr + Chr(34) + ") And LCase(Sf7) = LCase(dcr(" + Chr(34) + progEncr + Chr(34) + ")) Then LaLi = Supsabseee"
            
        'If cOd7 = dcr("ZvOE7QGSMHoaj38ZY2GONF59CZyn6q8+") And LCase(Sf7) = LCase(dcr("ZUPZP4Pq0ylE/pIFZgU24g==")) Then LaLi = Supsabseee
        Text1.Text = Text1.Text + renglon + vbCrLf + vbCrLf
    Next K
    
    
End Sub

Private Sub Form_Load()
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
End Sub

Private Sub List2_Click()
    If List2.ListIndex = -1 Then Exit Sub
        
    THIS = List2.List(List2.ListIndex)
    THIS = Split(THIS)(0)
    
    
    Set Fol2 = FSO.GetFolder(Fol1.Path + "\" + THIS)
    
    Dim FI As File
    For Each FI In Fol2.Files
        List1.AddItem Replace(FI.Name, ".txt", "")
    Next
    
End Sub
