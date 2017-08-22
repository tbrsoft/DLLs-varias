VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      TabIndex        =   15
      Text            =   "FH"
      Top             =   6630
      Width           =   855
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4170
      TabIndex        =   14
      Text            =   "SV"
      Top             =   6630
      Width           =   855
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3330
      TabIndex        =   13
      Text            =   "NR"
      Top             =   6630
      Width           =   855
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   12
      Text            =   "SFSFSF"
      Top             =   6630
      Width           =   3225
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   11
      Text            =   "fFINAL"
      Top             =   6210
      Width           =   7425
   End
   Begin VB.CommandButton Command6 
      Caption         =   "regenerar L37"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6210
      TabIndex        =   10
      Top             =   6630
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Extraer numero y ""final"" de un archivo de respuesta"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7590
      TabIndex        =   9
      Top             =   4530
      Width           =   4185
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   8
      Text            =   "idSoft_STR"
      Top             =   1140
      Width           =   1305
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Traducir LIC"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5280
      TabIndex        =   7
      Top             =   630
      Width           =   1425
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   5040
      Width           =   4185
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9570
      TabIndex        =   5
      Text            =   "66"
      Top             =   2340
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear Clave (con el numero de abajo y el mismo soft que tenia)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8640
      TabIndex        =   3
      Top             =   1590
      Width           =   3075
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1740
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver datos esta PC"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   450
      Width           =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   0
      X2              =   7740
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0442
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1965
      Left            =   7830
      TabIndex        =   4
      Top             =   120
      Width           =   3795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lee los datos de la PC y crea un archivo con todo el detalle (cosas unicas y no unicas)."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AP As String
Dim NF As New tbrNewSys.clsKEYS2
Dim FS As New Scripting.FileSystemObject

Dim TERR As New tbrErrores.clsTbrERR

Private Sub Command1_Click()
    On Local Error GoTo ERR4
    TERR.Anotar "aab"
    'este sistema obtiene datos unicos de la PC
    
    
    TERR.Anotar "aac"
    Dim ret As Long
    ret = NF.CFL(AP + "ifo.x", Text4.Text, VerInN, AP + "regLic.log")
    'graba un archivo encriptado con datos varios de la pc y los datos finales en el renglon "final"
    
    If ret > 0 Then TERR.AppendLog "dnw..1"
    'desencriptar y leer lo que se escribio.
    'esto es a titulo informativo para nosotros
    'NO SE USA EN EL SISTEMA
    TERR.Anotar "aad", ret
    ret = NF.Tasl(AP + "ifo.x", AP + "ifo.txt")
    If ret > 0 Then TERR.AppendLog "lilo..1"
    'esta funcion es para uso personal dentro del software no se usa. La usamos aca _
        para leer los datos que se toman
    
    'muestro estos datos
    TERR.Anotar "aae", ret
    Dim TE As TextStream
    Set TE = FS.OpenTextFile(AP + "ifo.txt", ForReading)
        Text1.Text = TE.ReadAll
    TE.Close
    
    'el archivo ifo.x es el que recibo yo y en base a ese puedo generar una clave que devuelve _
        'el final para verificar que se use en la misma pc y un numero de 1 a 1000 para otras _
        indicaciones
    TERR.Anotar "aaf"
    Exit Sub
    
ERR4:
    TERR.AppendLog "saaas"
    Resume Next
End Sub

Private Sub Command2_Click()
    On Local Error GoTo ERR4
    TERR.Anotar "aal2"
    Dim CM As New CommonDialog
    CM.DefaultExt = "L37"
    CM.DialogPrompt = "Indique donde se grabara la licencia"
    CM.DialogTitle = "Indique donde se grabara la licencia"
    
    CM.ShowSave
    Dim F As String
    F = CM.FileName
    TERR.Anotar "aam", F
    If F = "" Then Exit Sub
    
    TERR.Anotar "aan"
    If FS.FileExists(F) Then FS.DeleteFile F, True
    'no troco nada y por mas que lo haga el codigo del soft es el mismo del archivo que viene
    TERR.Anotar "aao"
    Dim ret As Long
    ret = NF.Desuso.LeerG(AP + "ifo.x", CLng(Text2.Text), F)
    If ret > 0 Then TERR.AppendLog "ck..1", CStr(ret)
    TERR.Anotar "aap", ret
    MsgBox "Se ha grabado " + vbCrLf + F + vbCrLf + "como respuesta de licencia"
    
    Exit Sub
ERR4:
    TERR.AppendLog "sbbb"
End Sub

Private Sub Command4_Click()
    On Local Error GoTo ERR4
    TERR.Anotar "aag"
    Dim CM As New CommonDialog
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    TERR.Anotar "aah", F
    If F = "" Then Exit Sub
    
    'lo copio aqui para hacer todo lo demas
    TERR.Anotar "aai"
    FS.CopyFile F, AP + "ifo.x"
    'desencriptar y leer lo que se escribio.
    'esto es a titulo informativo para nosotros
    'NO SE USA EN EL SISTEMA
    TERR.Anotar "aaj"
    Dim ret As Long
    ret = NF.Tasl(AP + "ifo.x", AP + "ifo.txt")
    If ret > 0 Then
        TERR.AppendLog "lilo..2"
        MsgBox "Error al traducir"
        Exit Sub
    End If
    'esta funcion es para uso personal dentro del software no se usa. La usamos aca _
        para leer los datos que se toman
    
    'muestro estos datos
    TERR.Anotar "aak", ret
    
    Dim TE As TextStream
    Set TE = FS.OpenTextFile(AP + "ifo.txt", ForReading)
        Text1.Text = TE.ReadAll
    TE.Close
    TERR.Anotar "aal"
    'el archivo ifo.x es el que recibo yo y en base a ese puedo generar una clave que devuelve _
        'el final para verificar que se use en la misma pc y un numero de 1 a 1000 para otras _
        indicaciones
    Exit Sub
ERR4:
    TERR.AppendLog "scccc"
    MsgBox Err.Description
End Sub

Private Sub Command5_Click()
    Dim CM As New CommonDialog
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    
    Dim FF2 As String
    Dim n_R As Long
    n_R = NF.Desuso.GetNR(F, FF2)
    Text3.Text = "Fue hecho para la PC: " + vbCrLf + FF2 + vbCrLf + vbCrLf + _
        "El numero usado fue: " + CStr(n_R) + vbCrLf + _
        "Para el software:" + NF.Desuso.GetSF + vbCrLf + _
        "Version del soft:" + CStr(NF.Desuso.getSV) + vbCrLf + _
        "Fecha de habilitacion:" + CStr(NF.Desuso.GetFH)
End Sub

Private Sub Command6_Click()
    On Local Error GoTo ERR4
    TERR.Anotar "aal2"
    Dim CM As New CommonDialog
    CM.DefaultExt = "L37"
    CM.DialogPrompt = "Indique donde se grabara la licencia"
    CM.DialogTitle = "Indique donde se grabara la licencia"
    
    CM.ShowSave
    Dim F As String
    F = CM.FileName
    TERR.Anotar "aam", F
    If F = "" Then Exit Sub
    
    TERR.Anotar "aan"
    If FS.FileExists(F) Then FS.DeleteFile F, True
    'no troco nada y por mas que lo haga el codigo del soft es el mismo del archivo que viene
    TERR.Anotar "aao"
    Dim ret As Long
    ret = NF.Desuso.LeerG("nofile" + _
        Chr(5) + Text5.Text + _
        Chr(5) + Text6.Text + _
        Chr(5) + Text8.Text + _
        Chr(5) + Text9.Text, CLng(Text7.Text), F)
    If ret > 0 Then TERR.AppendLog "ck..1", CStr(ret)
    TERR.Anotar "aap", ret
    MsgBox "Se ha grabado " + vbCrLf + F + vbCrLf + "como respuesta de licencia"
    
    Exit Sub
ERR4:
    TERR.AppendLog "sbbb"
End Sub

Private Sub Form_Load()
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    TERR.FileLog = AP + "regP2.log"
    TERR.Set_ADN CStr(App.Major * 100000 + App.Minor * 1000 + App.Revision)
    
    TERR.LargoAcumula = 600
    TERR.Anotar "aaa"
    
    NF.SetLog AP + "regKI.log"
    
    'definir los indices segun software
    SetIxE2Games
    
    TERR.AppendLog "INICIO LOG"
    
End Sub

'definir los indices segun soft
Private Sub SetIxE2Games()
    NF.setID 2, 1, 51
    NF.setID 2, 2, 23
    NF.setID 2, 3, 33
    NF.setID 2, 4, 14
    NF.setID 2, 5, 55
    
    NF.setID 3, 1, 63
    NF.setID 3, 2, 16
    NF.setID 3, 3, 21
    NF.setID 3, 4, 36
    NF.setID 3, 5, 66
    
    NF.setID 4, 1, 70
    NF.setID 4, 2, 49
    NF.setID 4, 3, 44
    NF.setID 4, 4, 74
    NF.setID 4, 5, 86
    
    NF.setID 5, 1, 83
    NF.setID 5, 2, 20
    NF.setID 5, 3, 46
    NF.setID 5, 4, 15
    NF.setID 5, 5, 82
    
    NF.setID 6, 1, 26
    NF.setID 6, 2, 81
    NF.setID 6, 3, 59
    NF.setID 6, 4, 23
    NF.setID 6, 5, 80
    
    NF.setID 7, 1, 4
    NF.setID 7, 2, 62
    NF.setID 7, 3, 77
    NF.setID 7, 4, 8
    NF.setID 7, 5, 31
    
End Sub

Private Function VerInN() As Long
    Dim L As Long
    L = App.Revision
    L = L + App.Minor * 1000
    L = L + App.Major * 100000
    
    VerInN = L
End Function
