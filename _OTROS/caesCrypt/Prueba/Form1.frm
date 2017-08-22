VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Prueba del Módulo Criptográfico de ca-es"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Proyecto realizado por http:\\www.escodigoabierto.com"
      Height          =   492
      Left            =   264
      TabIndex        =   16
      Top             =   5760
      Width           =   5364
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Manejo de archivos"
      Height          =   2700
      Index           =   1
      Left            =   72
      TabIndex        =   7
      Top             =   2832
      Width           =   6084
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Usar Base64 (Caracteres imprimibles)"
         Height          =   276
         Left            =   1896
         TabIndex        =   20
         Top             =   1680
         Width           =   3876
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1656
         Width           =   1692
      End
      Begin VB.CommandButton cmdEncriptarArchivo 
         Caption         =   "Encriptar Archivo"
         Height          =   420
         Left            =   120
         TabIndex        =   15
         Top             =   2112
         Width           =   5748
      End
      Begin VB.CommandButton txtVerArchivo 
         Caption         =   "Ver Archivo ENCRIPTADO"
         Height          =   348
         Left            =   1896
         TabIndex        =   14
         Top             =   720
         Width           =   3972
      End
      Begin VB.TextBox Text1 
         Height          =   324
         Index           =   4
         Left            =   1896
         TabIndex        =   9
         Text            =   "esta es la clave"
         Top             =   1152
         Width           =   3972
      End
      Begin VB.TextBox Text1 
         Height          =   324
         Index           =   3
         Left            =   1896
         TabIndex        =   8
         Text            =   "Archivo NO encriptado.txt"
         Top             =   288
         Width           =   3972
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Caption         =   "Llave (Key)"
         Height          =   252
         Index           =   5
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Caption         =   "Archivo encriptado:"
         Height          =   252
         Index           =   4
         Left            =   168
         TabIndex        =   11
         Top             =   792
         Width           =   1644
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Caption         =   "Archivo para encriptar"
         Height          =   252
         Index           =   3
         Left            =   0
         TabIndex        =   10
         Top             =   312
         Width           =   1812
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Manejo de Cadenas de Texto simples"
      Height          =   2604
      Index           =   0
      Left            =   72
      TabIndex        =   0
      Top             =   144
      Width           =   6084
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Usar Base64 (Caracteres imprimibles)"
         Height          =   276
         Left            =   1968
         TabIndex        =   19
         Top             =   1560
         Width           =   3876
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   216
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1560
         Width           =   1692
      End
      Begin VB.CommandButton cmdEncriptarString 
         Caption         =   "Encriptar String"
         Height          =   396
         Left            =   336
         TabIndex        =   13
         Top             =   1968
         Width           =   5508
      End
      Begin VB.TextBox Text1 
         Height          =   324
         Index           =   0
         Left            =   1968
         TabIndex        =   3
         Text            =   "Esta es la cadena que queremos encriptar"
         Top             =   288
         Width           =   3876
      End
      Begin VB.TextBox Text1 
         Height          =   324
         Index           =   1
         Left            =   1968
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   3876
      End
      Begin VB.TextBox Text1 
         Height          =   324
         Index           =   2
         Left            =   1968
         TabIndex        =   1
         Text            =   "esta es la clave"
         Top             =   1152
         Width           =   3876
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "String para encriptar:"
         Height          =   252
         Index           =   0
         Left            =   216
         TabIndex        =   6
         Top             =   312
         Width           =   1668
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "String ENCRIPTADO"
         Height          =   252
         Index           =   1
         Left            =   72
         TabIndex        =   5
         Top             =   792
         Width           =   1812
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "Llave (Key)"
         Height          =   252
         Index           =   2
         Left            =   1008
         TabIndex        =   4
         Top             =   1200
         Width           =   876
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oCripto As tbrCrypto.Crypt

Const SW_SHOW = 5

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub EncriptarString()
    
    Text1(1).Text = oCripto.EncryptString(Texto2Enum(Combo1.Text), Text1(0).Text, Text1(2).Text, IIf(Check1.Value = vbChecked, True, False))
    
End Sub

Private Sub EncriptarArchivo()
    
    Call oCripto.EncryptFile(Texto2Enum(Combo2.Text), App.Path & "\Archivo NO Encriptado.txt", App.Path & "\Archivo Encriptado.txt", Text1(4).Text, IIf(Check2.Value = vbChecked, True, False))
    
End Sub


Private Function Texto2Enum(ByVal Metodo As String) As eMetodoCriptografico
    Select Case Metodo
        
    Case "Blowfish"
        Texto2Enum = eMC_Blowfish
        
    Case "CryptAPI"
        Texto2Enum = eMC_CryptAPI
        
    Case "DES"
        Texto2Enum = eMC_DES
        
    Case "Gost"
        Texto2Enum = eMC_Gost
        
    Case "RC4"
        Texto2Enum = eMC_RC4
        
    Case "Skipjack"
        Texto2Enum = eMC_Skipjack
        
    Case "TEA"
        Texto2Enum = eMC_TEA
        
    Case "Twofish"
        Texto2Enum = eMC_Twofish
        
    Case "XOR"
        Texto2Enum = eMC_XOR
        
    End Select
    
End Function


Private Sub cmdEncriptarArchivo_Click()
    Set oCripto = New tbrCrypto.Crypt
    
    EncriptarArchivo
    
    Set oCripto = Nothing
    
End Sub

Private Sub cmdEncriptarString_Click()
    
    Set oCripto = New tbrCrypto.Crypt
    
    EncriptarString
    
    Set oCripto = Nothing
    
    
End Sub


Private Sub Command1_Click()
    Call ShellExecute(0, "Open", "http:\\www.escodigoabierto.com", "", App.Path, SW_SHOW)
End Sub

Private Sub Form_Load()
    'Carga los ComboBox
    With Combo1
        .AddItem "Blowfish"
        .AddItem "CryptAPI"
        .AddItem "DES"
        .AddItem "Gost"
        .AddItem "RC4"
        .AddItem "Skipjack"
        .AddItem "TEA"
        .AddItem "Twofish"
        .AddItem "XOR"
        
        .Text = "Blowfish"
    End With
    
    
    With Combo2
        .AddItem "Blowfish"
        .AddItem "CryptAPI"
        .AddItem "DES"
        .AddItem "Gost"
        .AddItem "RC4"
        .AddItem "Skipjack"
        .AddItem "TEA"
        .AddItem "Twofish"
        .AddItem "XOR"
        
        .Text = "Blowfish"
    End With
    
    
    Set oCripto = New Crypt
'    Inicializar Combo1.Text, oCripto
    EncriptarString
    
'    Inicializar Combo2.Text, oCripto
    EncriptarArchivo
    
    Set oCripto = Nothing
    
    
End Sub

Private Sub txtVerArchivo_Click()
    
    Call ShellExecute(0, "Open", App.Path & "\Archivo Encriptado.txt", "", App.Path, SW_SHOW)
    
End Sub


