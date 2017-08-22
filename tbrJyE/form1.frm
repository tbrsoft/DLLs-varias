VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00EBECE1&
   Caption         =   "Junta y Encripta"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00400000&
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar Lista"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9150
      TabIndex        =   11
      Top             =   150
      Width           =   1365
   End
   Begin VB.CheckBox chkSys32 
      BackColor       =   &H00EBECE1&
      Caption         =   "Copiar a sys32"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7950
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   5700
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar Como"
      Height          =   360
      Left            =   7950
      TabIndex        =   8
      Top             =   5250
      Width           =   2565
   End
   Begin VB.ComboBox cmbTipo 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   5250
      Width           =   3765
   End
   Begin VB.TextBox txtClave 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4050
      TabIndex        =   6
      Top             =   5250
      Width           =   3765
   End
   Begin VB.ListBox lstArchivos 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   150
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   450
      Width           =   10395
   End
   Begin VB.Label lblElemento 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "elementos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   150
      TabIndex        =   10
      Top             =   4350
      Width           =   750
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave de Encriptacion"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4050
      TabIndex        =   5
      Top             =   4950
      Width           =   2355
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Encriptacion"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   150
      TabIndex        =   4
      Top             =   4950
      Width           =   2235
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Elimine un archivo apretando Suprimir"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7650
      TabIndex        =   3
      Top             =   4800
      Width           =   2835
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Para agregar archivos arrastrelos hasta la lista."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6900
      TabIndex        =   2
      Top             =   4500
      Width           =   3615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Archivos"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Dim C As New tbrCrypto.Crypt
Dim juse As New tbrJUSE2.clsJUSE
Dim NombrePrograma As String
Dim WithEvents sEdit As EditorBase
Attribute sEdit.VB_VarHelpID = -1
Dim CD As New CommonDialog


Private Sub Command1_Click()
    If cmbTipo.ListIndex < 0 Then
        MsgBox "Elija un tipo de encriptacion valido!", vbCritical, NombrePrograma
        Exit Sub
    End If
    sEdit.GuardarComo
End Sub

Private Sub Command2_Click()
    lstArchivos.Clear
End Sub

Private Sub Form_Load()
    Set sEdit = New EditorBase
    NombrePrograma = "tbrJyE"
    
    cmbTipo.Clear
    cmbTipo.AddItem "Base 64"
    cmbTipo.AddItem "BlowFish"
    
    cmbTipo.AddItem "eMC_CryptAPI"
    cmbTipo.AddItem "eMC_DES"
    cmbTipo.AddItem "eMC_Gost"
    cmbTipo.AddItem "eMC_RC4"
    cmbTipo.AddItem "eMC_Skipjack"
    cmbTipo.AddItem "eMC_TEA"
    cmbTipo.AddItem "eMC_Twofish"
    cmbTipo.AddItem "eMC_XOR"
    
    
    sEdit.Iniciar "jus", NombrePrograma
    chkSys32.Value = Val(GetSetting("Editor", "Base", NombrePrograma + "_sys32", "0"))
    txtClave.Text = GetSetting("Editor", "Base", NombrePrograma + "_clave", "")
    cmbTipo.ListIndex = Val(GetSetting("Editor", "Base", NombrePrograma + "_combo", "0"))
End Sub

Private Sub lstArchivos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If lstArchivos.ListCount < 1 Then Exit Sub
        If lstArchivos.ListIndex < 0 Then Exit Sub
        lstArchivos.RemoveItem lstArchivos.ListIndex
        actualizar_lstelemetos
    End If
End Sub

Private Sub lstArchivos_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    For i = 1 To Data.Files.Count
        lstArchivos.AddItem Data.Files(i)
    Next i
    actualizar_lstelemetos
End Sub
Sub actualizar_lstelemetos()
    If lstArchivos.ListCount = 1 Then
        lblElemento = CStr(lstArchivos.ListCount) + " elemento"
    Else
        lblElemento = CStr(lstArchivos.ListCount) + " elementos"
    End If
End Sub


Private Sub sEdit_Grabar(elArchivo As String)
    Dim aux As String
    Dim s32 As String
    Dim i As Long
    
    'Unir los archivos
    juse.Archivo = elArchivo
    For i = 0 To lstArchivos.ListCount
        juse.AddFile lstArchivos.List(i)
    Next i
    juse.Unir
    
    
    'Recordar la configuracion
    SaveSetting "Editor", "Base", NombrePrograma + "_sys32", CStr(chkSys32.Value)
    SaveSetting "Editor", "Base", NombrePrograma + "_clave", txtClave.Text
    SaveSetting "Editor", "Base", NombrePrograma + "_combo", CStr(cmbTipo.ListIndex)
    
    'Encriptar
    EnciptarArchivo elArchivo, cmbTipo.ListIndex
    
    'Copiar a sys32
    If chkSys32.Value > 0 Then
        s32 = FSO.GetSpecialFolder(1)
        aux = Mid(elArchivo, InStrRev(elArchivo, "\"))
        aux = s32 + aux
        If Dir(aux) <> "" Then Kill aux
        FileCopy elArchivo, aux
    End If
    
    'Fin :)
    MsgBox "Archivo grabado exitosamente", vbInformation, NombrePrograma
End Sub

Sub EnciptarArchivo(qArchivo As String, cmbTipoIndex As Long)
    Dim aux As String
    Dim clave As String
    
    clave = txtClave.Text
    aux = qArchivo + "x"
    Select Case cmbTipoIndex
        Case 0
            ret = C.Base64File(qArchivo, aux, eB64_Encode)
        Case 1
            C.EncryptFile eMC_Blowfish, qArchivo, aux, clave
        Case Else
            MsgBox "metodo de encriptacion en desarrollo", vbInformation, NombrePrograma
            Exit Sub
    End Select
    ReemplazarArchivo qArchivo, aux
End Sub

Sub ReemplazarArchivo(QueArchivo As String, PorCualArchivo As String)
    Kill QueArchivo
    Name PorCualArchivo As QueArchivo
End Sub
