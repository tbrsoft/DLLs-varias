VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrWEBCAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'en resumen hay una iniciazión que es decirñle al driver que se conecte
'luego hay que elegir donde van las imagenes
'finalmente alcanzaria con usar la funcion de previsualizacion. En mi caso con la camara chota me jode
'otra opcion es usar get_frama y copy (al clipcboard) y despues meterlo en algun pictureox

Option Explicit

Public Enum evCamera 'eventos que puede atender el driver
    evVideoStream = 0 'este evento sucede antes GRABAR la imagen (FRAME). Aparentemente podriamos modificar los frames antes de grabarlos
    evOnFrame = 1 'igual que el anterior pero en vez de ser antes de GRABAR es antes de mostrar en pantalla!
    evStatus = 2 'saber el status, como esta todo
    evError = 3 'cuando llegue un error ir al procedimiento indicado
End Enum

Private sDriver As String
Private sVersion As String

'cuadros que tiene para mostrar, formularios propio del driver
Private DlgVideoDisplay As Boolean
Private DlgVideoFormat As Boolean
Private DlgVideoSource As Boolean
'otras capacidades
Private CapturaIniciadaOK As Boolean
Private DriverSoportaPaletas As Boolean
Private bHasOverlay As Boolean

Public Event Error(DetalleError As String, IdError As Long)
Public Event Status(DetalleStatus As String)

Private LWNDC As Long ' Handle to the Capture Windows
'para avisar de eventos uso esto
Private mHwndMsg As Long

Public Sub mHwndMsgSET(NEWmHwndMsg As Long)
    mHwndMsg = NEWmHwndMsg
    HwndMSGS = mHwndMsg
End Sub

Public Function GetDriverName() As String
    GetDriverName = sDriver
End Function

Public Function GetDriverVersion() As String
    GetDriverVersion = sVersion
End Function


'Friend Sub SetLwndc(newValue As Long)
'    LWNDC = newValue
'End Sub
'
'Friend Function GetLwndc(newValue As Long) As Long
'    GetLwndc = LWNDC
'End Function

Public Function InicioOK() As Boolean
    InicioOK = CapturaIniciadaOK
End Function

Public Function Puede_SoportarPaletas() As Boolean
    Puede_SoportarPaletas = DriverSoportaPaletas
End Function

Public Function Puede_Overlay() As Boolean
    Puede_Overlay = bHasOverlay
End Function

Public Function Puede_DlgVideoDisplay() As Boolean
    Puede_DlgVideoDisplay = DlgVideoDisplay
End Function

Public Function Puede_DlgVideoFormat() As Boolean
    Puede_DlgVideoFormat = DlgVideoFormat
End Function

Public Function Puede_DlgVideoSource() As Boolean
    Puede_DlgVideoSource = DlgVideoSource
End Function

Public Sub Mostrar_DlgVideoDisplay()
    If DlgVideoDisplay Then capDlgVideoDisplay LWNDC
    'vuelvo a acomodar ya que puede haber cambiado el tamaño
    ResizeCaptureWindow LWNDC
End Sub

Public Sub Mostrar_DlgVideoFormat_Click()
    If DlgVideoFormat Then capDlgVideoFormat LWNDC
    'vuelvo a acomodar ya que puede haber cambiado el tamaño
    ResizeCaptureWindow LWNDC
End Sub

Public Sub Mostrar_DlgVideoSource_Click()
    If DlgVideoSource Then capDlgVideoSource LWNDC
    'ResizeCaptureWindow lwndC
End Sub

Public Sub SacarFoto_ClipBoard()
    'meter en el portapapeles la imagen actual
    SendMessageS HwndMSGS, WM_SETTEXT, 0, "DLL Me piden Foto " + CStr(Timer)
    capGrabFrame LWNDC
    SendMessageS HwndMSGS, WM_SETTEXT, 0, "DLL Grabe foto" + CStr(Timer)
    capEditCopy LWNDC
    SendMessageS HwndMSGS, WM_SETTEXT, 0, "DLL copie clipboard" + CStr(Timer)
End Sub

Public Function IniciarEvento(Prender As Boolean, iniEvento As evCamera) As Long

    '**********************************************************
    'empezar o terminar de leer los eventos que necesito!
    'Aqui defino procedimientos que registran cosas durante el preview
    'asegurarse de al final desconectarlos!
    
    'devuelve lo que devuelva la funcion
    'por ejemplo el status no se puede prender si ya se esta grabando o previsualizando. Daria error
    
    If iniEvento = evError Then
        If Prender Then
            capSetCallbackOnError LWNDC, AddressOf MyErrorCallback
                'The WM_CAP_SET_CALLBACK_ERROR message sets an error callback function in the client _
                application. AVICap calls this procedure when errors occur.
        Else
            capSetCallbackOnError LWNDC, vbNull
        End If
    End If
    
    If iniEvento = evVideoStream Then
        If Prender Then
            capSetCallbackOnVideoStream LWNDC, AddressOf MyVideoStreamCallback
            'The WM_CAP_SET_CALLBACK_VIDEOSTREAM message sets a callback function in the application. _
                AVICap calls this procedure during streaming capture when a video buffer is filled.
            'The capture window calls the callback function before writing the captured frame to disk. _
                This allows applications to modify the frame if desired.
            'If a video stream callback function is used for streaming capture, the procedure must _
                be installed before starting the capture session and it must remain enabled for the _
                duration of the session. It can be disabled after streaming capture ends
        Else
            capSetCallbackOnVideoStream LWNDC, vbNull
        End If
    End If
    
    If iniEvento = evOnFrame Then
        If Prender Then
            capSetCallbackOnFrame LWNDC, AddressOf MyFrameCallback
            'The WM_CAP_SET_CALLBACK_FRAME message sets a preview callback function in the application. _
                AVICap calls this procedure when the capture window captures preview frames.
            'The capture window calls the callback function before displaying preview frames. This _
                allows an application to modify the frame if desired. This callback function is not _
                used during streaming video capture.
          
        Else
            capSetCallbackOnFrame LWNDC, vbNull
        End If
    End If
    
    If iniEvento = evStatus Then
        If Prender Then
            capSetCallbackOnStatus LWNDC, AddressOf MyStatusCallback
            'The WM_CAP_SET_CALLBACK_STATUS message sets a status callback function in the application. _
                AVICap calls this procedure whenever the capture window status changes
            'Applications can optionally set a status callback function. If set, AVICap calls this _
                procedure in the following situations:
                
                '* A capture session is completed.
                '* A capture driver successfully connected to a capture window.
                '* An optimal palette is created.
                '* The number of captured frames is reported.
          
        Else
            capSetCallbackOnStatus LWNDC, vbNull
        End If
    End If
    
End Function

Public Function Preview_Iniciar(Optional mRate As Integer = 50) As Long
    
    'que se ajuste a la imagen
    capPreviewScale LWNDC, True
    
    If Not capPreviewRate(LWNDC, mRate) Then
        Preview_Iniciar = 1
        'MsgBox "No se puede empezar a visualizar!"
        Exit Function
    End If

    'RaiseEvent Status("Set Rate Preview OK: " + CStr(mRate))
    'The WM_CAP_SET_PREVIEWRATE message sets the frame display rate in preview mode

    If Not capPreview(LWNDC, True) Then
        Preview_Iniciar = 2
        'MsgBox "No se puede enpezar a visualizar!"
        Exit Function
    End If
    
    'RaiseEvent Status("Empezo Preview OK")
    'The WM_CAP_SET_PREVIEW message enables or disables preview mode. In preview mode, _
        frames are transferred from the capture hardware to system memory and then _
        displayed in the capture window using GDI functions

    Preview_Iniciar = 0
    ResizeCaptureWindow LWNDC
End Function

Public Sub Preview_Detener()
    capPreview LWNDC, False
End Sub

Public Function GetDriverDescription()
    'devuelve un texto con lo que paso
    Dim lpszName As String * 255
    Dim lpszVer As String * 255
    
    'cargo la informacion del driver (?)
    capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
    'el primer parametro cero hace referencia a la primera camara plug y play instalada
    
    sDriver = lpszName
    sVersion = lpszVer
    
End Function

Public Function Driver_Conectar(hWndToShow As Long, Ancho As Long, Alto As Long) As Long
    
    'crear una ventana de captura (?)
      
    'LWNDC = capCreateCaptureWindowA(sDriver, WS_CHILD Or WS_VISIBLE, 0, 0, _
        Ancho, Alto, hWndToShow, 0)
        
    LWNDC = capCreateCaptureWindowA("tbrCamara", 0, 0, 0, _
        Ancho, Alto, hWndToShow, 0)
    'el ultimo parametro cero es el ID
        
    'conectar la ventana creada a la imagen que tengo. Imagino que el cero es el ultimo parametro
    'de la funciona anterior
    DoEvents
    'el siguiente devuelve falso si no se puede conectar el driver con la ventana creada
    If (capDriverConnect(LWNDC, 0)) Then
        Driver_Conectar = 0
    Else
        Driver_Conectar = 1 'error
    End If
        
End Function

Public Sub GetCapabilities()
    Dim Caps As CAPDRIVERCAPS
    
    'devuelve las capacidades del driver que estoy usando
    capDriverGetCaps LWNDC, VarPtr(Caps), Len(Caps)
    CapturaIniciadaOK = Caps.fCaptureInitialized
    DriverSoportaPaletas = Caps.fDriverSuppliesPalettes
    'si no puede crear paletas me parece que nada va a funcionar!
    DlgVideoDisplay = Caps.fHasDlgVideoDisplay
    DlgVideoFormat = Caps.fHasDlgVideoFormat
    DlgVideoSource = Caps.fHasDlgVideoSource
    bHasOverlay = Caps.fHasOverlay
End Sub

Public Sub Driver_Desconectar()
    DoEvents
    capSetCallbackOnError LWNDC, vbNull
    capSetCallbackOnStatus LWNDC, vbNull
    capSetCallbackOnYield LWNDC, vbNull
    capSetCallbackOnFrame LWNDC, vbNull
    capSetCallbackOnVideoStream LWNDC, vbNull
    capSetCallbackOnWaveStream LWNDC, vbNull
    capSetCallbackOnCapControl LWNDC, vbNull
End Sub
