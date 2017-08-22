VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "tbrIUdioma"
   ClientHeight    =   1590
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   450
      TabIndex        =   0
      Top             =   330
      Width           =   4485
   End
   Begin VB.Menu mnFile 
      Caption         =   "Archivo"
      Begin VB.Menu mnNewPhr 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnOpen 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mnSave 
         Caption         =   "Guardar"
      End
      Begin VB.Menu mnSaveAs 
         Caption         =   "Guarda como ..."
      End
      Begin VB.Menu mnSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnIdioma 
      Caption         =   "Idioma"
      Begin VB.Menu mnIDM0 
         Caption         =   "idm0"
         Index           =   0
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "Ayuda"
      Begin VB.Menu mnAbout 
         Caption         =   "Acerca de ..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    TERR.FileLog = AP + "regIDM.log"
    
    Me.Caption = "tbrSoft Idioma II v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    
    'TRADUCTOR DE ESTE MISMO PROGRAMA
    Set T = New tbrPrhase.clsQuickLang
    
    T.Load AP + "english.phr"
    'cargar la lista de idiomas en el menu
    Traducir
    
    'carga la lista de idiomas disponibles
    
    Dim Fi As File
    Dim Fo As Folder
    Set Fo = FSO.GetFolder(AP)
    Dim J As Long
    For Each Fi In Fo.Files
        If LCase(FSO.GetExtensionName(Fi.path)) = "phr" Then
            J = J + 1
            Load mnIDM0(J)
            mnIDM0(J).Caption = FSO.GetBaseName(Fi.path)
            mnIDM0(J).Visible = True
        End If
    Next
    mnIDM0(0).Visible = False
    
End Sub

Private Sub mnAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnIDM0_Click(Index As Integer)
    T.Load AP + mnIDM0(Index).Caption + ".phr"
    'cargar la lista de idiomas en el menu
    Traducir
End Sub

Private Sub mnNewPhr_Click()
    'crear uno nuevo casi vacio con lo idiomas basicos
    
    Dim C As New CommonDialog
    C.ShowSave
    
    Dim F As String
    F = C.FileName
    
    If F = "" Then Exit Sub
    
    Dim M As New tbrPrhase.clsPhraseMNG
    
    '-----------------------------------------------------
    Dim P As New tbrPrhase.clsPRHASE
    
    P.sID = "000001"
    P.BaseText = "Primera cadena"
    P.SetTrans "", "English"
    P.SetTrans "", "Portuges"
    M.AppendPHR P

    M.Save F
    
    Dim FRM As New frmTR
    FRM.OpenPhr M
    
End Sub

Private Sub mnOpen_Click()
    Dim C As New CommonDialog, F As String
    C.ShowOpen
    
    F = C.FileName
    
    If F = "" Then Exit Sub
    
    Dim M As New tbrPrhase.clsPhraseMNG
    
    M.Load F
    
    Dim FRM As New frmTR
    FRM.OpenPhr M
    
    
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub Traducir()
    mnFile.Caption = T.GetText("000001")
    mnNewPhr.Caption = T.GetText("000002")
    mnOpen.Caption = T.GetText("000003")
    mnSave.Caption = T.GetText("000004")
    mnSaveAs.Caption = T.GetText("000005")
    mnIdioma.Caption = T.GetText("000006")
    mnSalir.Caption = T.GetText("000007")
    mnHelp.Caption = T.GetText("000008")
    mnAbout.Caption = T.GetText("000009")
    Label1.Caption = T.GetText("000010", ";tbrSoft Internacional;Argentina")
End Sub
