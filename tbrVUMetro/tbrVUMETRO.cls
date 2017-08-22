VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrVUMETRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim AP As String

Public Event Pintar(D() As Byte)

Private WithEvents RELOJ As clsTimer
Attribute RELOJ.VB_VarHelpID = -1

Private Const MAXPNAMELEN = 32  '  longitud máx. del nombre del producto (incluido NULL)
Private Type WAVEOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
        dwSupport As Long
End Type

Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer    'wMid
    ProductID As Integer         'wPid
    DriverVersion As Long        'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_QUERY = &H1
Private Const WAVE_FORMAT_DIRECT = &H8
Private Const WAVE_FORMAT_DIRECT_QUERY = (WAVE_FORMAT_QUERY Or WAVE_FORMAT_DIRECT)
Private Const WAVE_FORMAT_PCM = 1   '  Necesario en archivos de recursos para #ifndef RC_INVOKED
Private Const WAVE_VALID = &H3         '  ;Interno


Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0
Private Const WAVE_MAPPER = -1&
'flags de WaveOutOpen
Private Const WAVE_MAPPED = &H4


Private Type WAVEFORMAT
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
End Type

Private Type MMTIME
        wType As Long
        U As Long
End Type

'funciones de WaveIn
Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private mIsPlaying As Boolean
Private mCurva As Double 'curvatura inicial de la elipse si se usara
Private CrecAnchoVU As Double 'crecimiento de los ovalos

Dim DevHandle As Long 'retorno de las llamadas a API
Dim LH As Long, RH As Long 'alto de las barra L y R (en cada momento)
Dim InData(0 To 511) As Byte 'todos los valores que entregan las apis sobre el sonido de cada momento
Dim m_Dispositivos As Long 'cantidad de placas (ver bien)
Dim m_inHabilitado As Boolean 'si no hay placas o hay errores
Dim BarrasEnVUmeter As Long 'bloques de colores totales
Dim m_CarFantastic As Boolean 'efecto auto fantastico. Se usa cuando no reproduce para mostrar algo cheto
Dim m_Borde As Long 'borde de los cuadros de colores
Dim m_EspacioEntreBarras As Long 'separacion entre cada barra
Dim contVu As Long 'mientras paso pintando es el numero del bloque de color actual
Dim contVuOff As Long 'lo mismo pero para los bloques apagados
Dim MaxLH As Long, MaxRH As Long 'ultimos valores maximos en cada canal
Dim ContTopVU As Long 'cuenta x vueltas manteniendo el ultimo maximo
Dim Caps() As WaveInCaps 'cada placa con todos sus valores
Dim tPics(1) As Object 'cajas de imagen donde se puede dibujar
Private Wave As WaveHdr

Public Property Set DefinePicture1(P1 As Object)
    Set tPics(0) = P1
End Property

Public Property Set DefinePicture2(P1 As Object)
    Set tPics(1) = P1
End Property


Public Property Get Dispositivos()
    Dim Which As Long
    
    For Which = 0 To waveInGetNumDevs - 1
        
        ReDim Preserve Caps(Which)
        Call waveInGetDevCaps(Which, VarPtr(Caps(Which)), Len(Caps(Which)))
        ''If Caps.Formats And WAVE_FORMAT_1M08 Then
        'if Caps.Formats And WAVE_FORMAT_1S08 Then _
            'Now is 1S08 -- Check for devices that can do stereo 8-bit 11kHz
        ''se cargan todos los dispositivos
            'Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
            'Devices(Which) = StrConv(Caps.ProductName, vbUnicode)
            'ccc.AddItem StrConv(Caps.ProductName, vbUnicode), Which
        'End If
    Next
    If Which = 0 Then
        'inhabilitar las barras
        m_inHabilitado = True
        'MsgBox "No tien salidas de audio!!!", vbCritical, "Ack!"
        Exit Property
    Else
        m_inHabilitado = False
    End If
    
    Dispositivos = Which
End Property

Private Sub RELOJ_Timer()

    Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
    Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
    Do
        'Nothing -- we're waiting for the audio driver to mark this wave chunk as done.
        DoEvents
    Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
    
    Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
    
    If DevHandle = 0 Then 'se cerro el dispositivo !
        RELOJ.Enabled = False
    End If
    DoEvents
    
    RaiseEvent Pintar(InData)
    'el minimo valor (no se porque) es 130, no he encontrado menos!
    'el maximo es 255 por que es una matriz de bytes
    
End Sub

Private Sub WriteFile(Arch As String, Texto As String)
    Dim Libre As Long
    Libre = FreeFile
    Open Arch For Output As Libre
        Write #Libre, Texto
    Close Libre
End Sub

Private Function ReadFile(Arch As String) As String
    Dim Libre As Long, Texto As String
    Libre = FreeFile
    Open Arch For Input As Libre
        Input #Libre, Texto
    Close Libre
    ReadFile = Texto
End Function

Public Property Let Curva(NewVal As Double)
    mCurva = NewVal
End Property

Public Property Get Curva() As Double
    Curva = mCurva
End Property

Public Property Get IsPlaying() As Boolean
    IsPlaying = mIsPlaying
End Property

Private Sub Class_Initialize()
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    'vu2***************************
    
    mCurva = 1.15
    CrecAnchoVU = mCurva
    'esto se inicia solo cuando se carga el control en ejecucuion
    'inicializar los dispositivos
    m_CarFantastic = False
    m_Borde = True
    m_EspacioEntreBarras = 25
    
    
    'vu2***************************
    
    
    m_Dispositivos = Dispositivos 'carga los dispositivos
    
    Set RELOJ = New clsTimer
    
    RELOJ.Enabled = False
    
End Sub

Public Sub SetIntervalLectura(i As Long)
    RELOJ.Interval = i
End Sub

Private Sub Class_Terminate()
    DoStop
End Sub

Property Get NombreDispositivo(Indice As Long) As String
    'solo lecture
    If Indice > UBound(Caps) Then
        NombreDispositivo = "No existe"
    Else
        NombreDispositivo = StrConv(Caps(Indice).ProductName, vbUnicode)
    End If
End Property

Property Let CarFantastic(new_CarFantastic As Boolean)
    m_CarFantastic = new_CarFantastic
    contVu = 0
End Property

Property Get CarFantastic() As Boolean
    CarFantastic = m_CarFantastic
End Property

Property Get inHabilitado() As Boolean
    'solo lectura
    inHabilitado = m_inHabilitado
End Property

Public Sub DoStop()
    RELOJ.Enabled = False
    'RELOJ.Interval = 0
    mIsPlaying = False
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
End Sub

Public Sub DoPause(Continuar As Boolean)
    RELOJ.Enabled = Continuar
End Sub

Public Function DoStart() As Long
    '0 todo ok
    '1 no se pudo abrir!

    DoStart = 0

    mIsPlaying = True
    'DoStop
    Static WAVEFORMAT As WaveFormatEx
    With WAVEFORMAT
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2 'Two channels -- left and right
        .SamplesPerSec = 11025
        .BitsPerSample = 8
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    Dim OpIN As Long
    
    OpIN = waveInOpen(DevHandle, 0, VarPtr(WAVEFORMAT), 0, 0, 0)
    'ver si ya esta abierto
    If OpIN = 4 Then
        'esta abierto. Cargar en DevHandle el valor de DevHandle devuelto en la primera apertura
        DevHandle = Val(ReadFile(AP + "devhw.dat"))
        'maercar como secundario
    End If
    If OpIN = 0 Then
        'se abrio ok debo grabar un archivo temporal para que lo lea este OCX cuando
        'se abra nuevamente
        WriteFile AP + "DevHw.dat", CStr(DevHandle)
    End If
        
    If DevHandle = 0 Then
        m_inHabilitado = True
        DoStart = 1
        'Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Function
    End If
    
    'si es secundario no vuelve a abrir nada
    waveInStart (DevHandle)
    m_inHabilitado = False
    
    'aparentemente aqui se define cual es la variable (matriz) que se a cargar con los datos
    Wave.lpData = VarPtr(InData(0))
    'el buffer podria ser mucho mas chico ya que yo solo uso el
    'InData(1) y el de InData(2)
    Wave.dwBufferLength = 512 'This is now 512 so there's still 256 samples per channel
    Wave.dwFlags = 0
    
    RELOJ.Enabled = True
    
End Function
