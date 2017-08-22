VERSION 5.00
Object = "{0371DBBE-C4D8-44B1-BFEE-712E91095894}#11.3#0"; "tbrListaConfig.ocx"
Begin VB.Form frmCFG2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   563
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   782
   StartUpPosition =   1  'CenterOwner
   Begin tbrListaConfig_CTL.ctlFullCFG CCF2 
      Height          =   1065
      Left            =   1080
      TabIndex        =   0
      Top             =   1335
      Width           =   2190
      _ExtentX        =   20399
      _ExtentY        =   14658
   End
End
Attribute VB_Name = "frmCFG2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AP As String

Private Sub CCF2_Fin(Rstrt As Long)
    If Rstrt = 0 Then
        'cerrar todo !
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Local Error GoTo errLD

    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    'CARGAR TODO

    Me.KeyPreview = True
        
    CCF2.SetFRMNEGRADA Me 'mnanu necesita un formulario
    CCF2.SetImgFondo App.Path + "\clsConfig\CONF.jpg"
    CCF2.SetPathImgConfs App.Path + "\PNG\"
    CCF2.SetSelConf App.Path + "\PNG\Selected.png"
    
    '/////////////////////////////
    'BASICO A TODOS (valores basicos que se pasan como parametros al pendrive)
    '/////////////////////////////
    CCF2.setPathError AP + "reg5.log"
    CCF2.SoftNow = "Test pulenta"
    CCF2.SetPcKy "IDtestPulentaaa"
    
    
    'DEFINIR LAS VARIABLES A USAR ANTES DE CARGAR! 'SEGUIRAQUI a medida que se usan cargarlas
    CCF2.VS.SetV "NombreSistema", "Test pulenta"
    CCF2.VS.SetV "ContadorHistorico", 17
    CCF2.VS.SetV "ContadorReiniciable", 99
    CCF2.VS.SetV "AP", AP
    
    'la lista de opciones del teclado es gigante !!!
    Dim listaKey As String
    listaKey = "Tecla A·65 && Tecla B·66 && Tecla C·67 && Tecla D·68 && Tecla E·69 && Tecla F·70 && Tecla G·71 && Tecla H·72 && Tecla I·73 && Tecla J·74 && Tecla K·75 && Tecla L·76 && Tecla M·77 && "
    listaKey = listaKey + "Tecla N·78 && Tecla O·79 && Tecla P·80 && Tecla Q·81 && Tecla R·82 && Tecla S·83 && Tecla T·84 && Tecla U·85 && Tecla V·86 && Tecla W·87 && Tecla X·88 && Tecla Y·89 && Tecla Z·90 && Tecla 0 (COMUN)·48 && "
    listaKey = listaKey + "Tecla 1 (COMUN)·49 && Tecla 2 (COMUN)·50 && Tecla 3 (COMUN)·51 && Tecla 4 (COMUN)·52 && Tecla 5 (COMUN)·53 && Tecla 6 (COMUN)·54 && Tecla 7 (COMUN)·55 && Tecla 8 (COMUN)·56 && Tecla 9 (COMUN)·57 && "
    listaKey = listaKey + "Tecla 0 (TECLADO NUMERICO)·96  && Tecla 1 (TECLADO NUMERICO)·97 && Tecla 2 (TECLADO NUMERICO)·98  && Tecla 3 (TECLADO NUMERICO)·99 && Tecla 4 (TECLADO NUMERICO)·100 && Tecla 5 (TECLADO NUMERICO)·101 && "
    listaKey = listaKey + "Tecla 6 (TECLADO NUMERICO)·102 && Tecla 7 (TECLADO NUMERICO)·103 && Tecla 8 (TECLADO NUMERICO)·104 && Tecla 9 (TECLADO NUMERICO)·105 && Tecla * (TECLADO NUMERICO)·106 && Tecla SIGNO MÁS (+) (TECLADO NUMERICO)·107 && "
    listaKey = listaKey + "Tecla INTRO (TECLADO NUMERICO)·108 && Tecla SIGNO MENOS (-) (TECLADO NUMERICO)·109 && Tecla PUNTO DECIMAL (.) (TECLADO NUMERICO)·110 && Tecla SIGNO DE DIVISIÓN (/) (TECLADO NUMERICO)·111 && Tecla F1·112 && "
    listaKey = listaKey + "Tecla F2·113 && Tecla F3·114 && Tecla F4·115 && Tecla F5·116 && Tecla F6·117 && Tecla F7·118 && Tecla F8·119 && Tecla F9·120 && Tecla F10·121 && Tecla F11·122 && Tecla F12·123 && Tecla F13·124 && Tecla F14·125 && "
    listaKey = listaKey + "Tecla F15·126 && Tecla F16·127 && Tecla SUPR·12 && Tecla ENTRAR·13 && Tecla MAYÚS·16 && Tecla CTRL·17 && Tecla MENÚ·18 && Tecla PAUSA·19 && Tecla BLOQ MAYÚS·20 && Tecla ESC·27 && Tecla BARRA ESPACIADORA·32 && "
    listaKey = listaKey + "Tecla RE PÁG·33 && Tecla AV PÁG·34 && Tecla FIN·35 && Tecla INICIO·36 && Tecla FLECHA IZQUIERDA·37 && Tecla FLECHA ARRIBA·38 && Tecla FLECHA DERECHA·39 && Tecla FLECHA ABAJO·40 && Tecla SELECT·41 && "
    listaKey = listaKey + "Tecla IMPRIMIR PANTALLA·42 && Tecla EXECUTE·43 && Tecla SNAPSHOT·44 && Tecla INS·45 && Tecla SUPR·46 && Tecla AYUDA·47 && Tecla BLOQ NUM·144"

    CCF2.VS.SetV "listaKey", listaKey
    
    Dim listaKeyH2K As String
    listaKeyH2K = "Señal 01·1 && Señal 02·2 && Señal 03·3 && Señal 04·4 && Señal 05·5 && Señal 06·6 && "
    listaKeyH2K = listaKeyH2K + "Señal 07·7 && Señal 08·8 && Señal 09·9 && Señal 10·10&& Señal 11·11&& "
    listaKeyH2K = listaKeyH2K + "Señal 12·12&& Señal 13·13&& Señal 14·14&& Señal 15·15&& Señal 16·16&& "
    listaKeyH2K = listaKeyH2K + "Señal 17·17&& Señal 18·18&& Señal 19·19&& Señal 20·20&& Señal 21·21&& "
    listaKeyH2K = listaKeyH2K + "Monedero 1 (M1)·22 && Monedero 2 (M1)·23"
    
    CCF2.VS.SetV "listaKeyh2k", listaKeyH2K
    
    'FINALMENTE CARGO LA CONFIG y LA MUESTRO
    
    'el archivo de configuracion esta encriptado en "skin.png", hay que desecriptarlo, cargarlo y borrarlo
    CCF2.Load AP + "config.txt"
    
    CCF2.GoRight
    
    Exit Sub
errLD:
    MsgBox Err.Description
    Resume Next
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    CCF2.Top = 4
    CCF2.Left = 4
    CCF2.Width = Me.Width - 8
    CCF2.Height = Me.Height - 24
End Sub

Private Sub TV_Click()
    If TV.SelectedItem Is Nothing Then Exit Sub
    lblAyuda.Text = TV.SelectedItem.Tag
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyZ
            CCF2.GoLeft
        Case vbKeyX
            CCF2.GoRight
        Case vbKeySpace
            CCF2.GoOK
        Case vbKeyEscape
            CCF2.GoCancel
    End Select
End Sub
