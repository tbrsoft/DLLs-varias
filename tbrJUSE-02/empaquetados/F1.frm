VERSION 5.00
Begin VB.Form F1 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "agregar carpeta"
      Enabled         =   0   'False
      Height          =   495
      Left            =   60
      TabIndex        =   6
      Top             =   2910
      Width           =   1365
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Extraer"
      Enabled         =   0   'False
      Height          =   495
      Left            =   60
      TabIndex        =   5
      Top             =   690
      Width           =   1365
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Empaquetar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   3420
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nuevo Js"
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   1890
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar suelto"
      Enabled         =   0   'False
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   2400
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "abri JS"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   150
      Width           =   1365
   End
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   1590
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   7245
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CM As New CommonDialog

Dim JSCreate As New tbrJUSE.clsJUSE
Dim JsExtract As New tbrJUSE.clsJUSE

Dim FSo As New Scripting.FileSystemObject

Dim LastFolderUsed As String

Private Sub Command1_Click()
    
    CM.DialogPrompt = "Elija el archivo a a abrir"
    CM.DialogTitle = "Elija el archivo a a abrir"
    CM.InitDir = LastFolderUsed
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    Dim res As Long
    res = JsExtract.ReadFile(F)
    If res = 1 Then
        MsgBox "No es una archivo empaquetado"
        Exit Sub
    End If
    
    List1.Clear
    Dim A As Long
    
    For A = 1 To JsExtract.CantArchs
        List1.AddItem JsExtract.GetListFiles(A, False)
    Next A
    
    'activar el extraer
    Command5.Enabled = True
    
    'desactiva agregar y empaquetar
    Command2.Enabled = False
    Command4.Enabled = False
    Command6.Enabled = False
    
End Sub

Private Sub Command2_Click()
    
    CM.InitDir = LastFolderUsed
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    If FSo.FileExists(F) Then JSCreate.AddFile F
    
    'mostrar la lista actualziada
    'ahora ya se podria empaquetar
    Command4.Enabled = True
    
    updateListCreate
End Sub

Private Sub updateListCreate()
    List1.Clear
    
    Dim A As Long
    For A = 1 To JSCreate.CantArchs
        List1.AddItem JSCreate.GetListFiles(A, True)
    Next A
    
End Sub

Private Sub Command3_Click()
    
    CM.InitDir = LastFolderUsed
    CM.DialogPrompt = "Nombre del archivo a crear"
    CM.DialogTitle = "Nombre del archivo a crear"
    CM.ShowSave
    
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    
    JSCreate.clearAll
    List1.Clear
    JSCreate.Archivo = F
    
    'activar el extraer
    Command5.Enabled = False
    
    'desactiva agregar y empaquetar
    Command2.Enabled = True
    Command4.Enabled = False 'todavia no tiene nada!
    Command6.Enabled = True
    
    
    updateListCreate
End Sub

Private Sub Command4_Click()
    JSCreate.Unir
    
    JSCreate.clearAll
    List1.Clear
    
    Command2.Enabled = False
    Command6.Enabled = False
    Command4.Enabled = False
    
    MsgBox "Se empaqueto ok"
End Sub

Private Sub Command5_Click()

    CM.DialogPrompt = "Elija el directorio a extraer"
    CM.DialogTitle = "Elija el directorio a extraer"
    
    CM.InitDir = LastFolderUsed
    
    CM.ShowFolder
    Dim F As String
    F = CM.InitDir
    If F = "" Then Exit Sub
    
    Dim A As Long
    For A = 1 To JsExtract.CantArchs
        JsExtract.Extract F, A
    Next A
    
    MsgBox "Se extrajeron correctamente " + CStr(JsExtract.CantArchs) + " archivos"
    
End Sub

Private Sub Command6_Click()
    CM.InitDir = LastFolderUsed
    CM.ShowFolder
    Dim F As String
    F = CM.InitDir
    
    If F = "" Then Exit Sub
    'agregar uno por uno todos los archivos de la carpeta
    Dim FI As File, FO As Folder
    Set FO = FSo.GetFolder(F)
    For Each FI In FO.Files
        JSCreate.AddFile FI.path
    Next
    'mostrar la lista actualziada
    'ahora ya se podria empaquetar
    Command4.Enabled = True
    
    updateListCreate
End Sub

Private Sub Form_Load()
    LastFolderUsed = App.path
End Sub
