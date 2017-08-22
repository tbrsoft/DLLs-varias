VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desencriptar Karaokes"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   554
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00B3FF00&
      Height          =   285
      Left            =   4440
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   5
      Top             =   4800
      Width           =   2400
   End
   Begin VB.TextBox CSali 
      BackColor       =   &H00B3FF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "C:\"
      Top             =   5190
      Width           =   6750
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00B3FF00&
      Caption         =   "Directorio de Salida"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4710
      Width           =   2205
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00B3FF00&
      Caption         =   "Desencriptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   6930
      Picture         =   "Form1.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4620
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00B3FF00&
      Caption         =   "Borrar Lista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6810
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   1395
   End
   Begin VB.ListBox Caja 
      BackColor       =   &H00B3FF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   3660
      ItemData        =   "Form1.frx":0894
      Left            =   120
      List            =   "Form1.frx":0896
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   540
      Width           =   8085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desencriptador de MN1 a MN0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   210
      TabIndex        =   6
      Top             =   60
      Width           =   3405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CD As New CommonDialog

Dim TE As New tbrSoftEncrMan.clstbrENC
Private CDK_prefix(6) As String 'prefijos sabidos para cada cd existente
Private CDK_qey(6) As String 'clave que existe para cada prefijo

Private Sub Caja_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For I = 1 To Data.Files.Count
        Dim arch As String
        arch = Data.Files(I)
        If LCase(Right(arch, 3)) = "mn1" Then
            Caja.AddItem arch
        End If
    Next I
End Sub

Private Sub Command1_Click()
Caja.Clear
End Sub

Private Sub Command2_Click()
Dim NeoNom As String

Dim N1 As String
Dim N2 As String

For I = 0 To Caja.ListCount - 1
    N1 = Caja.List(I)
    'ver que clave tiene segun su prefijo
    'probar uno por uno los CDs existentes
    Dim KYY As String, PX As String 'clave,prefijo encontrados
    KYY = GetH(N1, PX)
    
    If KYY = "NOIDENTIFICOCD" Then
        'este no pertenece a ningun cd oficial de tbrSoft de karaoke
        Caja.List(I) = "MALO:" + Caja.List(I)
        GoTo SIG
    End If
    
    NeoNom = CSali + Mid(N1, InStrRev(N1, "\") + 1)
    
    NeoNom = Mid(NeoNom, 1, Len(NeoNom) - 3) + "mn0"
    
    TE.Encriptar True, KYY, N1, NeoNom, PX
    Caja.List(I) = "OK:  " + Caja.List(I)
SIG:
Next I

MsgBox "Listo!"
End Sub

Private Sub Command3_Click()
    CD.InitDir = ""
    CD.DialogTitle = "Carpeta de Archivos Disponibles"
    CD.ShowFolder
    
    If CD.InitDir <> "" Then
        CSali = CD.InitDir
    End If

    If Right(CSali, 1) <> "\" Then CSali = CSali + "\"
    
End Sub

Private Sub Form_Load()
    'para usar karaokes
    CDK_prefix(0) = "asjdfsadfsadfsadfsadfsadfasdfsa546456465"
    CDK_qey(0) = "sdfuoyhsdfsdiufyaoisfSAD789F6AD78F6A7SD89F6A89S6F879AS"
    
    CDK_prefix(1) = "rrweqwrwerwerrrrrrrrr23423r223r2r23r2r32r23r2r23r"
    CDK_qey(1) = "yyssysyasuoisdyoa8sdy8a9dsysa978dsyaasrea98"
     
    CDK_prefix(2) = "fuigwsyfs7idfs8d6f9a8s76d879as6f987as6df876879fas6d987"
    CDK_qey(2) = "sdfystdf78we6f9872r6798wyefuihwdjfhw8euyr3279hiuwgfiwegfiywegfo78"
    
    CDK_prefix(3) = "sdfysuftas6df7asdtf6a8s76f"
    CDK_qey(3) = "sadfsoiudfyws98efyw987ef69weyf789w6fy978wgfe8wyef879wyt8"
    
    CDK_prefix(4) = "sdf78sydf8s7gf8sctys87dcyt8s7ycdsy7sd8cy7s"
    CDK_qey(4) = "sdvcuyhsdgbv8ywetgv76wetf76wetf67wtec76wstc76ewt76etc76wect67wetc867w"
    
    CDK_prefix(5) = "asdfsa9d8f7sa98fda7d87qw6dq987wd879qwd97q8d9w87q6d987q6wd98ss"
    CDK_qey(5) = "asdfiuyadais7ydta7sdt78qw6tdq6w8td6qwdq6wtd9wq6td98q76d7qtw78dtq78wdt89q"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Hace_Click()
    If Hace.Value = 0 Then
        Hace.Caption = "Desencriptar"
    Else
        Hace.Caption = "Encriptar"
    End If
End Sub

Private Function GetH(AR As String, PX As String) As String
    'devuelve la clave para abrirlo
    'solo se ingresa el archivo MN1
    'en el parametro PX devuelve el prefijo encontrado
    
    Dim KKY As String
    KKY = "NOIDENTIFICOCD"
    PX = "NO"
    
    Dim J As Long
    Dim resPX As String
    
    For J = 0 To 6 'pruebo todos los cds posibles
        resPX = GetPrefixKar(AR, Len(CDK_prefix(J)))
        If resPX = CDK_prefix(J) Then
            'si tiene la licencia
            KKY = CDK_qey(J)
            PX = CDK_prefix(J)
            Exit For
        End If
    Next J
    
    GetH = KKY
End Function

Public Function GetPrefixKar(AR As String, nLEN As Long) As String
    'devuelve los primeros caracteres de un karaoke encriptado (AR) para saber
    'con clave abrirlo
    Dim Aux As String
    
    If Dir(AR) = "" Then Exit Function
    Aux = Space(nLEN)
    Open AR For Binary As #1
        Get #1, 1, Aux
    Close #1
    
    GetPrefixKar = Aux
End Function

