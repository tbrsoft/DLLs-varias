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
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   4590
      TabIndex        =   17
      Top             =   1320
      Width           =   2595
   End
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   4530
      Width           =   4185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Extraer numero y ""final"" de ESTE archivo de respuesta"
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
      Top             =   4080
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
      Caption         =   "Ver datos otra PC"
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
'Este es un ejemplo de como usar el sistema de licencias que tenemos
'es una DLL que puede obtener datos especificos de un equipo que no cambian
'cuando se formatea y resinstala ni cuando cambia el sistema operativo
'o sea nos permite identificar una PC
'cambia cuando:
'se reemplaza alguno de los siguientes componentes
'disco rigido
'procesador
'placas de red (instalar o agregar nuevas, activar o desactivas desde la bios)
'placa madre
'si se actualzia la bios tambien cambia

'la funcion  DoNow graba un archivo encriptado con estos datos.
'Graba datos que cambian y un dato
'final con el resumen de los elementos que no cambian
'Esta funcion (DoNow) se ejecuta del lado del cliente cuando se necesite enviar a tbrSoft
'una vez hecho el doNow se carga la funcion GetRF que devuelve solo los valores finales
'este RF se usa para comparar cuando llegue el archivo de licencia que enviamos nosotros

'una vez que nos envian el archivo que genera doNow nosotros usamos la funcion CK
'NF.CK AP + "ifo.x", CLng(Text2.Text), AP + "res_ifo.lic"
'el primer parametro es el archivo que recibimos, el segundo es un numero del 1 a 100 que
'puede representar el tipo de licencia o lo que se nos ocurra
'el ultimo parametro es un archivo que enviaremos de respuesta
'este contiene encritado el codigo de la pc en el que funcionara y el numero

'cuando el cliente recibe la clave debe hacer la comparacion con la funcion GetNR
'Dim n_R As Long
'n_R = NF.GetNR(AP + "res_ifo.lic", FF2)

'getNR necesita como parametro el archivo de licencia
'FF2 es un parametro byRef que devuelve la cadena final del archivo que se lee
'la funcion devuelve el numero que creo tbrSoft al licenciar

'entonces cuando el usuario ingresa un archivo de licencia se debe hacer la comparacion

'If lcase(NF.GetRF) = lcase(FF2) Then
'    Select Case n_R
'        Case 17, 44, 33, 22 'es un licencia tipo 1
'            MsgBox "Se ha insertado licencia tipo 1"
'        Case 91, 92, 3 'es una licencia tipo 2
'            MsgBox "Se ha insertado licencia tipo 2"
'        Case Else 'no la cree yo para este programa
'            MsgBox "Error en la licencia"
'    End Select
'
'End If

'se recomienda no usar esto al principio del sistema como para que sea facil de interceptar.
'los programas deben arrancar ok. Pueden usar algun modulo que administre este sistema de la
'manera que les guste. O como mínimo una variable que guarde el resultado ya que hacer todo
'esto cada vez que necesiten saber que licencia tiene es pesado



Option Explicit
Dim AP As String
Dim NF 'As New tbrDATA.clsTODO
Dim FS 'As New Scripting.FileSystemObject
Dim TERR 'As New tbrErrores.clsTbrERR

Private Sub INIT()
    'MsgBox "1"
    Set TERR = New tbrErrores.clsTbrERR
    'MsgBox "2"
    Set FS = New Scripting.FileSystemObject
    'MsgBox "3"
    'Dim CC As New tbrCrypto.Crypt
    'MsgBox "3.5"
    Set NF = New tbrDATA.clsTODO
    'MsgBox "4"
End Sub

Private Sub Command1_Click()
    On Local Error GoTo ERR4
    TERR.Anotar "aab"
    'este sistema obtiene datos unicos de la PC
    NF.SetSF Text4.Text 'nuevo!! defino para que soft es!!!
    TERR.Anotar "aac"
    Dim ret As Long
    ret = NF.DoNow(AP + "ifo.x") 'graba un archivo encriptado con datos varios de la pc y _
        los datos finales en el renglon "final"
    If ret > 0 Then TERR.AppendLog "dnw..1"
    'desencriptar y leer lo que se escribio.
    'esto es a titulo informativo para nosotros
    'NO SE USA EN EL SISTEMA
    TERR.Anotar "aad", ret
    ret = NF.LiToLo(AP + "ifo.x", AP + "ifo.txt")
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
    ret = NF.CK(AP + "ifo.x", CLng(Text2.Text), F)
    If ret > 0 Then TERR.AppendLog "ck..1", CStr(ret)
    TERR.Anotar "aap", ret
    MsgBox "Se ha grabado " + vbCrLf + F + vbCrLf + "como respuesta de licencia"
    
    Exit Sub
ERR4:
    TERR.AppendLog "sbbb"
End Sub

Private Sub Command3_Click()
    Dim FF2 As String
    Dim n_R As Long
    n_R = NF.GetNR(AP + "res_ifo.lic", FF2)
    Text3.Text = "Fue hecho para la PC: " + vbCrLf + FF2 + vbCrLf + vbCrLf + _
        "El numero usado fue: " + CStr(n_R) + vbCrLf + "Para el software:" + NF.GetSF
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
    
    'ver si quiere con otra clave de encriptacion NUEVO DIC 2010
    If Text10.Text <> "" Then
        NF.kkyNew Text10.Text
    End If
    
    Dim ret As Long
    ret = NF.LiToLo(AP + "ifo.x", AP + "ifo.txt")
    If ret > 0 Then
        TERR.AppendLog "lilo..2"
        Text1.Text = "error al abrir el archivo=" + CStr(ret)
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
End Sub

Private Sub Command5_Click()
    Dim CM As New CommonDialog
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    
    Dim FF2 As String
    Dim n_R As Long
    n_R = NF.GetNR(F, FF2)
    Text3.Text = "Fue hecho para la PC: " + vbCrLf + FF2 + vbCrLf + vbCrLf + _
        "El numero usado fue: " + CStr(n_R) + vbCrLf + _
        "Para el software:" + NF.GetSF + vbCrLf + _
        "Version del soft:" + CStr(NF.getSV) + vbCrLf + _
        "Fecha de habilitacion:" + CStr(NF.GetFH)
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
    ret = NF.CK("nofile" + _
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
    
    INIT
    
    TERR.FileLog = AP + "regP2.log"
    TERR.Set_ADN CStr(App.Major * 100000 + App.Minor * 1000 + App.Revision)
    
    TERR.LargoAcumula = 600
    TERR.Anotar "aaa"
    
    NF.SetLog AP + "regKI.log"
    
    TERR.AppendLog "INICIO LOG"
    
End Sub
