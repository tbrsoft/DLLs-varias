VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   4455
      Left            =   60
      ScaleHeight     =   4395
      ScaleWidth      =   7215
      TabIndex        =   9
      Top             =   4860
      Width           =   7275
      Begin VB.CommandButton Command9 
         Caption         =   "cCred"
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   2820
         Width           =   1275
      End
      Begin VB.TextBox tLog 
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   23
         Top             =   900
         Width           =   7035
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Encender"
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Apagar"
         Height          =   255
         Left            =   6060
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   6975
         TabIndex        =   17
         Top             =   300
         Width           =   7035
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   645
         End
         Begin VB.Label lEst 
            AutoSize        =   -1  'True
            Caption         =   "en espera"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   720
            TabIndex        =   18
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   6300
         Top             =   2640
      End
      Begin MSWinsockLib.Winsock WS 
         Left            =   6720
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox picCFG 
         Height          =   1095
         Left            =   60
         ScaleHeight     =   1035
         ScaleWidth      =   7035
         TabIndex        =   11
         Top             =   3240
         Width           =   7095
         Begin VB.TextBox tPuerto 
            Height          =   285
            Left            =   6240
            TabIndex        =   14
            Text            =   "8881"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox tPath 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Text            =   "c:\"
            Top             =   360
            Width           =   6855
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Guardar Configuracion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5460
            TabIndex        =   12
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Puerto"
            Height          =   195
            Left            =   5640
            TabIndex        =   16
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Path de Licencias"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   1470
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Conexiones multpiples simultaneas estan prohibidas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   2685
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servicio Online"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   1050
      End
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1260
      TabIndex        =   8
      Top             =   300
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtOBS 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   4770
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.ComboBox cmbST 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4380
      TabIndex        =   5
      Top             =   3720
      Width           =   1425
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5910
      TabIndex        =   4
      Top             =   3720
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   3720
      Width           =   1425
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   7335
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Licencias ya creadas"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   1
      Top             =   900
      Width           =   2955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'tbrSLW=========================
Dim CheckCrd As New clsCliManager

Dim WithEvents SLW As clsSLW
Attribute SLW.VB_VarHelpID = -1
Dim PathLic As String
Dim Puerto As Long

Private Type GenerateMSG
    MSG As String
    Path As String
End Type
'===============================

'Para activar licencias del licenciador: Ctrl+Alt+Shift+F6
Option Explicit
Dim AP As String
Dim sFileBase As String 'archivo con la definicion de uso de este licenciaero
Dim Folder1138 As String 'carpeta de las licencias generadas (NO INCLUYE \ FINAL !!!)

Dim NF As New tbrNewSys.clsKEYS2 ' licencias que se activan de e2Games
Dim NF2 As New tbrNewSys.clsKEYS2 ' licencia de este licenciero
'si este sistema no tiene licencias solo generar licencias gratuitas

Dim FS As New Scripting.FileSystemObject
Dim TERR As New tbrErrores.clsTbrERR
Dim FSO As New Scripting.FileSystemObject

Dim ct As Long 'conteo de licencias para saber cuantas usa

Dim OBS() As String 'observaciones de cada generacion de licencias

'generarme una licencia para mi mismo licenciero
Private Sub Command5_Click()
    
    TERR.Anotar "bah-44", VerInN
    
    On Local Error GoTo ERR4
    
    Dim ret As Long
    Dim F As String
    F = AP + NF.Desuso2.sName  'el archivo LIC es en realidad un archivo con el nombre del sistema que este licenciewro va a licenciar
    TERR.Anotar "aam-44", F
    
    Dim F2 As String 'destino de la lic, el mismo nombre con otra extension
    F2 = F + ".L38"
    
    TERR.Anotar "aam22-44", F2
    
    'me aseguro que no este la licencia para crearla de nuevo
    If FS.FileExists(F2) Then FS.DeleteFile F2, True
    TERR.Anotar "aam25-44"
    
    'definir que mnr se le dara a la licencia (tener en cuenta que necesita segun archivo LOAD)
    Dim nrUse As Long
    'esta es en realidad una pregunta engañosa, generalmente debe escribirse 7!
    nrUse = CLng(InputBox("Cuantas licencias desea programar ?", , 100)) 'licencia para el licenciero, 7 es SL pero algunos soft requieren 5!!
    
    
    TERR.Anotar "aam27", nrUse
    ret = NF2.Desuso.LeerG(F, nrUse, F2) 'crear el L38 correspondiente segun el NR ingresado
    TERR.Anotar "baj", ret, nrUse
    
    If ret > 0 Then
        If ret = 1 Then
            TERR.AppendLog "Leerg-bak-1"
            'MsgBox "El archivo recibido no parece un archivo válido"
            MsgBox getID2("aliOhQSuX/8mhWUcWTUi7a1LF6dHxwS02MUtsRHtXy19Zn37ljCYZU0wTpHF5O7Q6dHolMjdQfN9/gNY1Flfbw==")
        Else
            TERR.AppendLog "Leerg-bak", CStr(ret)
            'MsgBox "Error al crear licencia (02-" + CStr(ret) + ")" + vbCrLf + "Envie informe a tbrSoft"
            MsgBox getID2("/gkUn9SvKjCsJZ/2JaYufjBryGS+F5qit+bBoD9ZJ7zugKetVwWCWw==") + _
                CStr(ret) + ")" + vbCrLf + getID2("h8jM68Bt+Ueyn8QUDkYAw8IGVFGRRFbUODDv2+A/q5LkJ5vV1QJThQ==")
        End If
    Else
        UpdateGL2 'que se vea si se cargo ok
        TERR.AppendSinHist "res__3_8=" + F2 + CStr(NF2.GL)
    End If
    
    Exit Sub
    
ERR4:
    TERR.AppendLog "sbbb", TERR.ErrToTXT(Err)
    'MsgBox "Error al generar licencia, envie el registro de errores a tbrSoft"
    MsgBox getID2("1HNxtJug6YhP9m1ilXVnNrdH3TQUkAkJdtc19x3ziNkuvl18+3aJxnqlfO5HK37OMH9vATczsbyCeHsRxT+f5WU6TOjn2oMvFB3+yg+NYAg=")
End Sub

Private Sub Command9_Click()
    Dim cCred As Long
    Dim IdUsuario As Long
    Dim IdSoft As Long
    
    IdUsuario = 145
    IdSoft = 1
    
    cCred = CheckCrd.GetCredCli(IdUsuario, IdSoft)
    MsgBox cCred
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'ANTES DE ESTYO ELEGIR QUE SISTEMA SE LICENCIARA
    'para activar esto hay que apretar Ctrl+Alt+Shift+F6
    
    If KeyCode = vbKeyF6 And Shift = 7 Then
    'If KeyCode = vbKeyF6 Then
        Command5.Visible = True
    End If
End Sub

Private Sub Form_Load()
    On Local Error GoTo errINI
    
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    TERR.FileLog = AP + "regP2.log"
    TERR.Set_ADN CStr(VerInN)
    
    TERR.LargoAcumula = 600
    TERR.Anotar "aaa"
    
    sFileBase = AP + "base.unl"
    
    'este de verdad es el registro de errores
    NF.SetLog AP + "regL01.log"
    NF2.SetLog AP + "regL02.log"
    
    'Crear nueva licencia
    Command2.Caption = getID2("U9YmQGGOS17MdKnLoQARHc8Xoj2Q+EYTTEFfnJLfRQ4=")
    
    'Regenerar licencia elegida
    Command1.Caption = getID2("qmY+5LeiiHtsLYrfodJyzIvIcq7dzS/9BOKJ3DDUWc8D66zQq2E31A==")
    
    'Generar informe errores
    Command4.Caption = getID2("09bFTG71vkZz3/egRdFnzcPvn2lRPG/e1ar7whnOsbai3nnwl95Bnw==")
    
    'Salir
    Command3.Caption = getID2("ERSoIWg3kXgElkrsegGjDZSnGfI+LhLE")
    
    'como voy a poner teclas especiales necesito que el frm las agarre
    Me.KeyPreview = True
    
    'cargar el archivo con datos
    UnLoad__C
    
    'tbrSLW===========================================================================
    Set SLW = New clsSLW
    'Leo las configuraciones
    PathLic = GetSetting("tbrSLW_gui", "cfg", "PathLicencia", App.Path + "\Licencias\")
    Puerto = Val(GetSetting("tbrSLW_gui", "cfg", "Puerto", "8881"))
    tPath = PathLic
    tPuerto = CStr(Puerto)
    
    If SLW.InicializarPath(PathLic) = 1 Then
        'El directorio no existe NI SE PUEDE CREAR
        'MsgBox "El Path para las licencias no existe!" + vbCrLf + _
            "El programa no va a funcionar, cambie el Path por uno existente" _
            , vbInformation, "Path no existente"
            
        PathLic = App.Path + "\Licencias\"
        tPath = PathLic
        SLW.InicializarPath (PathLic)
    End If
    SLW.InicializarSocket WS
    
    'Iniciar Automaticamente
    Command8_Click
    '=================================================================================
    
    
    
    Exit Sub
errINI:
    TERR.AppendLog "ini091", TERR.ErrToTXT(Err)
    Resume Next
End Sub

'el load basico
Private Function UnLoad__C()
    
    On Local Error GoTo ErrLoadPhrase
    
    'ver si hay solo un archivo para cargar o varios!!
    Dim FI As File
    Dim FO As Folder
    Set FO = FSO.GetFolder(AP + "loads")
    
    cmbST.Clear: cmbST.Visible = False 'combo de los SofTwares
    
    For Each FI In FO.Files 'que se carguen todos los archivos con la extension correspondiente
        If LCase(FSO.GetExtensionName(FI.Path)) = "load" Then cmbST.AddItem FI.Name
    Next
    
    If cmbST.ListCount = 0 Then 'si no hay  nada dar la posibilidad de cargar
         'no le doy oportunidad de cargar otro
'        'seguiraqui ERROR, NO SE QUE DEFINICION HAY!!1
'        Dim CM As New CommonDialog
'        CM.InitDir = AP
'        CM.ShowOpen
'
'        If CM.FileName <> "" Then
'            sFileBase = CM.FileName
'            Unload__D sFileBase 'CARGARLO!!!!
'        Else
            TERR.AppendLog "NoUNL"
            Exit Function
'        End If
    End If
    
    If cmbST.ListCount = 1 Then 'si es uno solo es ese sin cambiar
        Unload__D AP + cmbST.List(0)
    End If
    
    If cmbST.ListCount > 1 Then 'si son varios dejarlo elegir
        cmbST.Visible = True
        
        'agregar para licenciarse a si mismo
        
    End If
    
    'tbrSLW: Seleccion 3pm automaticamente
    cmbST.ListIndex = 6
    
    
    Exit Function
    
ErrLoadPhrase:
    TERR.AppendLog "Unload__C__Err", TERR.ErrToTXT(Err)
    
End Function


'regenerar licencia elegida
Private Sub Command1_Click()
    If List1.ListIndex = -1 Then Exit Sub
    
    'ver cual esta elegido
    Dim Sp() As String
    Sp = Split(List1)
    
    Dim FIL As String
    FIL = Folder1138 + "\" + Sp(1) + ".L38"
    
    If FSO.FileExists(FIL) Then
        Dim CM As New CommonDialog
        'CM.DialogPrompt = "Elija el destino del archivo"
        CM.DialogPrompt = getID2("Ezms2ha1OHwzqx9ngwN4z5mMPvQFPWXbl3BnEap6e5hXHV0ggr+ggw==")
        CM.DialogTitle = getID2("Ezms2ha1OHwzqx9ngwN4z5mMPvQFPWXbl3BnEap6e5hXHV0ggr+ggw==")
        
        CM.InitDir = AP
        
        CM.ShowFolder
        
        If CM.InitDir = "" Then Exit Sub
        Dim FI As String
        FI = FSO.BuildPath(CM.InitDir, FSO.GetBaseName(FIL) + ".L38")
        
        FSO.CopyFile FIL, FI
        
        'MsgBox "Se ha copiado con exito"
        MsgBox getID2("G1/NLkPX9TqqXopjaaEMvQMRVcWNJyzaPewVaTobabcudmxwMqMIzw==")
    Else
        'MsgBox "No se encontro la licencia buscada"
        MsgBox getID2("nvYDIFQ6XeEv+PmWAf/t3ingTpa27IGdNMcmgGalXpeO4mxg07SrOdJeeMc9+fEc")
    End If
    
End Sub

'generar licencia
'Esta funcion devuelve dos valores String
'MSG=Es la resputesta... Si funciono, hubo un error, etc.
'Path es el path de la lic generada
Private Function Generar(F As String, PAD As String) As GenerateMSG
    'F=Archivo Licencia
    'PAD=detalles de aquin registra licencia
    Dim ret As String
    
    TERR.Anotar "bah", VerInN
    UpdateGL2 'revalidar la licencia de este generador de licencias
    
    On Local Error GoTo ERR4
    
    'seguiraqui
    'revisar todas las licencias anteriores y ver si ya se habia creado licencia para esa pc
    'no lo muestro por que puede delatar fallas en licencias sin MACAdress
    
    'moverlo a la carperta lic! pero recordar su ubicacion original para que la licencia vaya alli!
    Dim fORIG As String
    fORIG = F
    
    Generar.Path = ""
    
    'evito que se usen nombres de archivos repetidos
    Dim CambiaNombre As Boolean
    
    Dim DestF As String
    DestF = FSO.BuildPath(Folder1138, FSO.GetBaseName(F))
    
    'me pasa que f y destf van al mismo archivo pero uno con formato 8.3 y el otro no..
    'no se me ocurre otra forma de compararlo
    Dim Fi1 As File
    Dim Fi2 As File
    
    Set Fi1 = FSO.GetFile(F)
    If FSO.FileExists(DestF) Then
        Set Fi2 = FSO.GetFile(DestF)
    
        'si son archivos distintos le busco un nuevo nombre
        If Fi1.ShortPath + Fi1.ShortName <> Fi2.ShortPath + Fi2.ShortName Then
            CambiaNombre = GetNewNameIfNeed(DestF, 1)
            TERR.Anotar "aam20", CambiaNombre, DestF
            If CambiaNombre Then
                'avisarle
                'MsgBox "Ya habia un pedido de licencia con ese nombre, se reemplazo el nombre por:" + vbCrLf + DestF
                Generar.MSG = "Ya habia un pedido de licencia con ese nombre, se reemplazo el nombre por:" + vbCrLf + DestF
                
            End If
        End If
    Else
        FSO.CopyFile F, DestF
    End If
    
    Dim F2 As String 'destino de la lic, el mismo nombre con otra extension
    F2 = DestF + ".L38"
    
    Dim F3 As String 'datos sobre la activacion de la licencia fecha # de la de licencia
    F3 = DestF + ".dat"
    
    TERR.Anotar "aam22", F2
    
    'me aseguro que no este la licencia
    If FS.FileExists(F2) Then FS.DeleteFile F2, True
    TERR.Anotar "aam25"
    If FS.FileExists(F3) Then FS.DeleteFile F3, True
    TERR.Anotar "aam26"
    'no toco nada y por mas que lo haga el codigo del soft es el mismo del archivo que viene
    TERR.Anotar "aao"
    
    Dim nrUse As Long
    nrUse = useNR 'si este sistema tiene licencia activa lic premium y si no gratuitas
    TERR.Anotar "aam27", nrUse
    ret = NF.Desuso.LeerG(F, nrUse, F2) 'crear el L38 correspondiente
    TERR.Anotar "baj", ret, nrUse
    
    If ret > 0 Then
        If ret = 1 Then
            TERR.AppendLog "Leerg-bak-1"
            'MsgBox "El archivo recibido no parece un archivo válido"
            'MsgBox getID2("aliOhQSuX/8mhWUcWTUi7a1LF6dHxwS02MUtsRHtXy19Zn37ljCYZU0wTpHF5O7Q6dHolMjdQfN9/gNY1Flfbw==")
            Generar.MSG = getID2("aliOhQSuX/8mhWUcWTUi7a1LF6dHxwS02MUtsRHtXy19Zn37ljCYZU0wTpHF5O7Q6dHolMjdQfN9/gNY1Flfbw==")
        Else
            TERR.AppendLog "Leerg-bak", CStr(ret)
            'MsgBox "Error al crear licencia (02-" + CStr(ret) + ")" + vbCrLf + "Envie informe a tbrSoft"
            'MsgBox getID2("/gkUn9SvKjCsJZ/2JaYufjBryGS+F5qit+bBoD9ZJ7zugKetVwWCWw==") + _
                CStr(ret) + ")" + vbCrLf + getID2("h8jM68Bt+Ueyn8QUDkYAw8IGVFGRRFbUODDv2+A/q5LkJ5vV1QJThQ==")
            Generar.MSG = getID2("/gkUn9SvKjCsJZ/2JaYufjBryGS+F5qit+bBoD9ZJ7zugKetVwWCWw==") + _
                CStr(ret) + ")" + vbCrLf + getID2("h8jM68Bt+Ueyn8QUDkYAw8IGVFGRRFbUODDv2+A/q5LkJ5vV1QJThQ==")
        End If
    Else
        TERR.Anotar "aap", ret
        'grabar el registro (pidiendo datos de a quien va)
        
        Dim r As String
        r = CStr(ct) + Chr(5) + NF.Desuso.GFlic(F) + Chr(5) + FSO.GetBaseName(DestF) + Chr(5) + PAD 'este ultimo es lo que se va a cargar en los renglones de la lista de licencias, si dejo el F de antes los archivos renombrados apareceran con su nombre anterior!
        
        Dim TE As TextStream
        Set TE = FSO.CreateTextFile(F3, True)
            TE.Write r
        TE.Close
        ct = ct + 1 'sumar al contador!
        
        TERR.Anotar "aap5"
        
        'que se grabe el ap+LIC para recuperar y tambien en el origen
        FSO.CopyFile F2, fORIG + ".L38", True
        
        'MsgBox "Se ha grabado " + vbCrLf + F2 + vbCrLf + "como respuesta de licencia"
        'MsgBox getID2("Q0xlWklwcrKkGIb4EGAt8M38KpJXlWtRYDLPYcrrtiU=") + vbCrLf + fORIG + ".L38" + vbCrLf + _
            getID2("F4jN41N+/HWiGF3ld2Ne+rgepuFY3MpUo97ZbrF9czZP8nm8+5NvIA==")
        Generar.MSG = getID2("Q0xlWklwcrKkGIb4EGAt8M38KpJXlWtRYDLPYcrrtiU=") + vbCrLf + fORIG + ".L38" + vbCrLf + _
            getID2("F4jN41N+/HWiGF3ld2Ne+rgepuFY3MpUo97ZbrF9czZP8nm8+5NvIA==")
        
        Generar.Path = fORIG + ".L38"

        
        TERR.AppendSinHist "res__3_8=" + F2
        
        UPL
    End If
    
    Exit Function
    
ERR4:
    TERR.AppendLog "sbbb", TERR.ErrToTXT(Err)
    'MsgBox "Error al generar licencia, envie el registro de errores a tbrSoft"
    'MsgBox getID2("1HNxtJug6YhP9m1ilXVnNrdH3TQUkAkJdtc19x3ziNkuvl18+3aJxnqlfO5HK37OMH9vATczsbyCeHsRxT+f5WU6TOjn2oMvFB3+yg+NYAg=")
    Generar.MSG = getID2("1HNxtJug6YhP9m1ilXVnNrdH3TQUkAkJdtc19x3ziNkuvl18+3aJxnqlfO5HK37OMH9vATczsbyCeHsRxT+f5WU6TOjn2oMvFB3+yg+NYAg=")
End Function

Private Sub Command2_Click()
    Dim ret As Long
    TERR.Anotar "aal2"
    Dim CM As New CommonDialog
    '"Inserte archivo de licencia" = getid2("YsY2soqPiS8k63O2xsEbGLjdw9hpSkTp3KuP6kt5eyMN6vJ8IEvDxg==")
    CM.DialogPrompt = getID2("YsY2soqPiS8k63O2xsEbGLjdw9hpSkTp3KuP6kt5eyMN6vJ8IEvDxg==")
    CM.DialogTitle = getID2("YsY2soqPiS8k63O2xsEbGLjdw9hpSkTp3KuP6kt5eyMN6vJ8IEvDxg==")
    
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    TERR.Anotar "aam", F
    
    If F = "" Then Exit Sub
    
    Dim PAD As String
    PAD = InputBox("Indique detalles de aquin registra licencia para futuras busquedas", , "No deje en blanco!")

    MsgBox Generar(F, PAD).MSG
End Sub

'busca un nuevo nombre para un archivo, lo cambia de pecho y da true si tuvo que cambiarlo

Private Function GetNewNameIfNeed(ByRef sFile As String, Optional DoCopy As Long = 0) As Boolean
    
    'me puede pedir que lo copie con doCopy=1
    'o que lo mueva con DoCopy=2
    'o nada en cero
    
    If FSO.FileExists(sFile) = False Then
        GetNewNameIfNeed = False
        Exit Function
    Else
        Dim FOL As String
        Dim ftEST As String, fExt As String
        Dim U As Long
        FOL = FSO.GetParentFolderName(sFile)
        ftEST = FSO.GetFileName(sFile)
        fExt = FSO.GetExtensionName(sFile)
        
        Dim RES As String
        U = 2
        Do
            TERR.Anotar "aan", U, sFile
            'compatible con archivos sin extencion
            If fExt <> "" Then
                RES = FSO.BuildPath(FOL, ftEST + "__" + CStr(U) + "." + fExt)
            Else
                RES = FSO.BuildPath(FOL, ftEST + "__" + CStr(U))
            End If
            
            If FSO.FileExists(RES) = False Then Exit Do
            U = U + 1
        Loop
        
        If DoCopy = 1 Then
            FSO.CopyFile sFile, RES
        End If
        
        If DoCopy = 2 Then
            FSO.MoveFile sFile, RES
        End If
        
        sFile = RES
        GetNewNameIfNeed = True
    End If
End Function

'volver a crear el archivo de licencia LIC de este generador de licencias
'y volver a cargar el archivo de licencia!
Private Sub UpdateGL2()
    
    On Local Error GoTo ERR4
    TERR.Anotar "bak"
    'generar el archivo de licencia (LIC) de esta pc (tiene que devolver cero!)
    'se refiere a la licencia de este licenciero
    
    Dim ret As Long
    ret = NF2.CFL(AP + NF.Desuso2.sName, NF.Desuso2.sName, VerInN, AP + "regL02.log")
    TERR.Anotar "bal", ret
    
    If ret > 0 Then
        TERR.AppendLog "Leerg-bam", CStr(ret)
        'MsgBox "Error al crear licencia (01-" + CStr(ret) + ")" + vbCrLf + "Envie informe a tbrSoft"
        MsgBox getID2("cHhE5NGDKnnYr4yXcPKs43r0fjXqh/nVFcp1462KwCWJcz4o82vAnQ==") + _
            CStr(ret) + ")" + vbCrLf + getID2("l02mNKQvUaRyXjjo31cat5+MGsVIGWesZjuIBXbEKFC5Rn2FaJyc0g==")
    End If
    
    'revisar si este licenciero tiene licencia para poder saber si puede generar licencias pulentas o solo licencias gratuitas
    NF2.IFL AP + NF.Desuso2.sName + ".L38", NF.Desuso2.sName, VerInN, True
    
    'aqui ya queda cargada GL que es el valor de la licencia del licenciero que se consulta despues
    
    TERR.Anotar "bal"
    'veo si es licencia pulenta
    Dim J1 As Long, J2 As Long
    Randomize
    J1 = CLng(Rnd * 80) + 10
    Randomize
    J2 = CLng(Rnd * 80) + 10
    'la funcion GL devuelve el numero de licencia, por ejemplo 7 cuando es SL
    TERR.Anotar "bai", CStr(J1) + CStr(Year(Date) - 2000 + NF2.GL) + CStr(J2)
    
    'Generador de licencias L38
    Me.Caption = getID2("ut+ZUSHIRRyGrJ17h+uenHwT8RrhxrTli97efLfek+CHLbfPgdwbuA==") + CStr(NF2.GL) + NF.Desuso2.sName
    
    Exit Sub
    
ERR4:
    TERR.AppendLog "err0293jmm", TERR.ErrToTXT(Err)
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

'generar informe de errores
Private Sub Command4_Click()

    On Local Error GoTo ErrJSA
    
    TERR.Anotar "baa"
    Dim J As New tbrJUSE.clsJUSE
    Dim fileINFOR As String
    
    fileINFOR = AP + "Informe" + "." + CStr(Year(Date)) + "." + CStr(Month(Date)) + "." + CStr(Day(Date)) + "." + CStr(Hour(Time)) + "." + CStr(Minute(Time)) + ".JSA"
    
    Dim FilsJSA(2) As String
    FilsJSA(0) = AP + "regP2.log"
    FilsJSA(1) = AP + "regL01.log"
    FilsJSA(2) = AP + "regL02.log"
    
    J.clearAll
    
    J.Archivo = fileINFOR
    
    Dim K As Long
    For K = 0 To UBound(FilsJSA)
        J.AddFile FilsJSA(K)
    Next K
    
    Dim ret As Long
    ret = J.Unir
    
    If ret <> 0 Then
        TERR.AppendLog "bae", "RET=" + CStr(ret)
        MsgBox "Error al crear informe !" + vbCrLf + "(une:" + CStr(ret) + ")"
    Else
        MsgBox "El informe esta listo para enviarse en:" + vbCrLf + _
            fileINFOR
        TERR.Anotar "bac"
        
        For K = 0 To UBound(FilsJSA)
            TERR.Anotar "bad", K
            If FSO.FileExists(FilsJSA(K)) Then FSO.DeleteFile FilsJSA(K), True
        Next K
        
        
    End If
    
    J.clearAll
    
    Exit Sub
    
ErrJSA:
    TERR.AppendLog "bab", TERR.ErrToTXT(Err)
    MsgBox "Error al crear informe !" + vbCrLf + "(fin:" + CStr(Err.Number) + ")"
End Sub

'si este programa no tiene licencia las licencias que generar son gratuitas
'si esta licenciados sacara licencias full!
Private Function useNR()
    'hay muchas opciones al pedo que no son posibles pero no quiero me crackeen esto !
    Dim t As Long, T2 As Long
    t = NF2.GL 'licencia del licenciero
    
    If t > 0 Then 'No existen indices negativos para las matrices!!
        T2 = NF.Desuso2.GetComoDev(t)
    Else
        T2 = NF.Desuso2.GetComoDev(0)  'PARA TODOS LOS NEGATIVOS !!
    End If
    
    TERR.Anotar "bag", t, T2
    
    useNR = T2
End Function

Private Function VerInN() As Long
    Dim L As Long
    L = App.Revision
    L = L + App.Minor * 1000
    L = L + App.Major * 100000
    
    VerInN = L
End Function

'es un dcr encriptador
Private Function getID2(if2 As String) As String
    Dim fr As New tbrCrypto.Crypt
    Dim fr2 As String
    fr2 = fr.DecryptString(eMC_Blowfish, if2, "ID invalido genere uno nuevo", True)
    getID2 = fr2
End Function

Public Function STRceros(n As Variant, Cifras As Integer) As String
    'n es el numero y cifras es la cantidad final de cifras del str terminado
    'devuelve ej : para 232,6 = 000232 para 1902,12 = 000000001902
    'complaeta con ceroas adelante
    ' si n es mas lasgo que cifras devuelve el valor n sin ningun cero adelante
    Dim STRn As String
    STRn = Trim(CStr(n))
    Dim DIF As Integer
    DIF = Cifras - Len(STRn)
    If DIF > 0 Then
        Dim CEROstr As String
        CEROstr = String(DIF, "0")
        STRceros = CEROstr + STRn
    Else
        STRceros = STRn
    End If
End Function

'cargar todo lo del elegido
Private Sub cmbST_Click()
    Unload__D AP + "loads\" + cmbST
End Sub

'QUEDEAQUI
'se empezo a pasar a la DLL y no se limpio aqui ni se aplico !!
'24 11 09
'cargar un archivo de datos de software que ya existe
Private Sub Unload__D(sFile__D As String)

    On Local Error GoTo ErrLoadPhrase
    TERR.Anotar "baf", sFile__D
        
    sFileBase = sFile__D
    Dim ret As Long
    
    'cargar los indices del archivo load en cada objeto de licencia
    ret = NF.SetIXs(sFileBase, 1)
    TERR.Anotar "baj", ret
    
    'este NF2 es el propio del licenciero
    ret = NF2.SetIXs(sFileBase, 2)
    TERR.Anotar "baj2", ret
    
    'se asegura que exista ap+LIC
    Folder1138 = AP + "LIC" + NF.Desuso2.sName
    If FSO.FolderExists(Folder1138) = False Then FSO.CreateFolder Folder1138 'carpeta donde pongo todos los archivos
    
    'QUEDEAQUI queda el ultimo, no hace 2 procesos independientes !!!
    
    UPL 'mostrar todas las existentes
    ct = ct + 1 'el maximo mas uno
    
    'para que actualice el caption del form
    UpdateGL2
    
    Exit Sub
    
ErrLoadPhrase:
    TERR.AppendLog "Unload__D__Err", TERR.ErrToTXT(Err)
End Sub

'actualizar la lista de licencias
Private Sub UPL()

    On Local Error GoTo errUPL
    List1.Clear
    Dim foLI As String
    foLI = Folder1138
    'ver todas las licencias activadas
    If FSO.FolderExists(foLI) = False Then FSO.CreateFolder foLI
    Dim FI As File
    Dim FO As Folder
    
    ct = 0 'contador de mnumero de licencia
    
    Set FO = FSO.GetFolder(foLI)
    For Each FI In FO.Files
        If FSO.GetExtensionName(FI.Path) = "dat" Then
            'abrirlo, controlar el numero de licencia y sumarle uno para controlar
            Dim TE As TextStream, RR As String
            Set TE = FSO.OpenTextFile(FI.Path, ForReading, False)
                If TE.AtEndOfStream Then
                    TERR.AppendLog "aaa32", FI.Path
                    RR = ""
                Else
                    RR = TE.ReadAll
                End If
            TE.Close
            
            If RR <> "" Then
                'ver el numero de licencia y guardar el mayor
                Dim Sp() As String, thisCT As Long
                Sp = Split(RR, Chr(5))
                If IsNumeric(Sp(0)) Then
                    thisCT = CLng(Sp(0)) + 1
                    
                    'es un renglon pulenta agregarlo a la lista!!
                    List1.AddItem STRceros(thisCT, 5) + " " + Sp(2)
                    
                    'recordar observaciones
                    ReDim Preserve OBS(List1.ListCount)
                    If UBound(Sp) > 2 Then
                        OBS(List1.ListCount) = Sp(3)
                    End If
                    
                Else
                    thisCT = 0
                    TERR.AppendLog "aaa33", RR
                End If
                'si es un numero de orden de licencia mayor lo guardo como el ultimo
                If thisCT > ct Then ct = thisCT
            End If
        End If
    Next
    
    Exit Sub
    
errUPL:
    TERR.AppendLog "aaa98", TERR.ErrToTXT(Err)
    Resume Next
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    
    
    
    List1.Width = Me.Width - 250
    List1.Left = 45
    List1.Height = Me.Height / 3.6 '- List1.Top - Command1.Height - 510
    
    'Command1.Top = Me.Height - Command1.Height - 450
    Command1.Top = List1.Top + List1.Height + 50
    Command3.Top = Command1.Top
    Command4.Top = Command1.Top
    
End Sub

Private Sub List1_Click()
    If List1.ListIndex > -1 Then
        txtOBS.Text = OBS(List1.ListIndex + 1)
        txtOBS.Visible = True
        
    Else
        txtOBS.Visible = False
    End If
End Sub

Private Sub SLW_ArchivoRecibido(IdUsuario As Long, PathFile As String)
    Dim cCred As Long
    Dim IdSoft As Long
    
    IdSoft = 1
    
    
    'Tiene creditos suficientes?
    cCred = CheckCrd.GetCredCli(IdUsuario, IdSoft)
    
    If cCred > 0 Then
        'SI tiene creditos:
        
        'Activo la Licencia
        '??????????????
        
        'Resto un Credito
        CheckCrd.RestarCredito IdUsuario, IdSoft
        EnviarSuceso "Credito Descontado: " + CStr(cCred) + " creditos ahora"
        
        'Activar Lic
        Dim ret As GenerateMSG
        Dim PAD As String
        Dim nombreLicA As String
        
        PAD = "Lic" + CStr(Date) + "_" + CStr(IdUsuario)
        ret = Generar(PathFile, PAD)
        EnviarSuceso "Respuesta del Activador de Licencias: " + ret.MSG
        
        nombreLicA = Mid(ret.Path, InStrRev(ret.Path, "\") + 1)
        SLW.EnviarLicenciaActivada ret.Path, 1, nombreLicA
        

    Else
        'NO tiene creditos:
        SLW.EnviarCreditosInsuficientes
        
        EnviarSuceso "Creditos Insuficientes: " + CStr(cCred) + " creditos"
    End If
    'Restar
    
End Sub

'============================================================================
'tbrSLW: AQUI COMIENZA EL CODIGO DE tbrSLW ON-LINE
'============================================================================
Private Sub Timer1_Timer()
    lEst = SLW.GetStrEstado
End Sub

Private Sub SLW_Suceso(Suceso As String)
    tLog.Text = tLog.Text + Suceso + vbCrLf
End Sub

Private Sub Command6_Click()
    PathLic = tPath
    Puerto = Val(tPuerto)
    
    SaveSetting "tbrSLW_gui", "cfg", "PathLicencia", PathLic
    SaveSetting "tbrSLW_gui", "cfg", "Puerto", tPuerto
End Sub

Private Sub Command7_Click()
    SLW.DetenerServicio
End Sub

Private Sub Command8_Click()
    SLW.ComenzarServicio Puerto
End Sub

Private Sub EnviarSuceso(Suceso As String)
    Dim Codigo As Long
    Dim aux As String
    
    Codigo = 0
    aux = "[" + CStr(Date) + "]"
    aux = aux + "[" + CStr(Time) + "]"
    aux = aux + "[" + Format(Codigo, "000") + "]"
    aux = aux + ":"
    aux = aux + Suceso

    tLog.Text = tLog.Text + aux + vbCrLf

End Sub

'========================================================
'(!) Importante Conectar estos 2 eventos!
'========================================================
Private Sub WS_ConnectionRequest(ByVal requestID As Long)
    SLW.ConnectionRequest requestID
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
    SLW.DataArrival bytesTotal
End Sub

'============================================================================
'tbrSLW: FIN
'============================================================================

