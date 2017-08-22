VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTbrDx8ESP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'traductor y agrupador para hacer un funcionamiento mas transparente

Option Explicit

'**********************************
'objetos externos
Implements IDSOutCallback 'esta solo dentro del modulo SounOut en la DLL, _
    manda los eventos de cambio de estado y otros

Dim WithEvents TmrPlay As clsTimer
Attribute TmrPlay.VB_VarHelpID = -1

'**********************************
'placas de sonido, nombre, escripcion y un ID como referencia
Private clsMonoLoader As tbrDX8.MonotonLoader
'....traducido por tbr en:
Private Type PlacaSonido
    sID As String
    Nombre As String
    Descripcion As String
    Iniciada As Long '0 es no; 1 es si
End Type
'lista con las placas
Dim PA() As PlacaSonido
Dim mTotalPlacas As Long 'cantidad contada al inicialziar
Dim sArchivo As String 'path completo del archivo ejecutado
'****************************************
'para abrir y reproducir archivos
'****************************************
'archivo que se ejecuta. pueden ser hasta 4 al mismo tiempo
Private clsOutput(3)    As tbrDX8.SoundOut
' abre el archivo y es parametro para que se inicialize con sndOUT
Private clsStream(3)    As tbrDX8.ISoundStream
Private clsTag(3)       As tbrDX8.StreamTag
Private isOpen(3) As Long '0=cerrado, 1=abierto

'****************************************
'eventos definicion
'****************************************
Public Event CancionEmpeiza(IdSonido As Long)
Public Event CancionCorriendo(Index As Long, IdSonido As Long, strTime As String, MiliSecPlayed As Long, MiliSecToEnd As Long)
Public Event CancionTermina(IdSonido As Long)
Public Event CancionCambiaEstado(Index As Long, IdSonido As Long, NuevoEstado As String)
Public Event CancionSample(IdSonido As Long)

'****************************************
'fin lista de eventos
'****************************************
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const FFT_SAMPLES As Long = 512

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Private Sub Class_Initialize()
    mTotalPlacas = -1 'indica que nada se inicializo
    Set TmrPlay = New clsTimer
    TmrPlay.Enabled = False
    TmrPlay.Interval = 23 'varias por segundo
End Sub


Public Function LeerPlacas() As Long 'devuelve el numero de placas
'deja cargado los valores basicos de cada placa
    
    Dim i As Long
    Set clsMonoLoader = New tbrDX8.MonotonLoader
    'giro por cada placa leyendo sus valores
    For i = 1 To clsMonoLoader.DeviceCount
        ReDim Preserve PA(i)
        PA(i).Descripcion = clsMonoLoader.DeviceDescription(i)
        PA(i).Nombre = clsMonoLoader.DeviceName(i)
        'si el id es todo ceros es otra cosa... averiguar
        PA(i).sID = clsMonoLoader.DeviceGUID(i)
    Next
    mTotalPlacas = clsMonoLoader.DeviceCount
    LeerPlacas = mTotalPlacas
End Function

Public Function InicializarPlaca(mSampleRate As Long, iCanales As Integer, _
    iBitsPorSample As Integer, sID As String) As Long
    
    'iniciliza la placa elegida
    'ver si la placa arranca con valores pedidos para ver hasta donde soporta
    
    'DEVUELVE:
    '0: Todo OK
    '1:no se inicializo
    
    If Not clsMonoLoader.Initialize(mSampleRate, iCanales, iBitsPorSample, sID) Then
        InicializarPlaca = 0 'error no inicio
        PA(Get_Placa_IndexFromSID(sID)).Iniciada = 0 'ver que indice tiene ese ID
    Else
        InicializarPlaca = 1 'todo ok
        PA(Get_Placa_IndexFromSID(sID)).Iniciada = 1 'ver que indice tiene ese ID
    End If
End Function

Public Function Player_AbrirMP3(ssArchivo As String, indicePlayer As Long) As Long
    sArchivo = ssArchivo
    '0: todo OK
    '1: Falla en OpenSource (abrir el archivo)
    '2: falla al inicializar el archivo
    '3: no se puede ejecutar
    
    'limpieza anterior!
    If Not clsStream(indicePlayer) Is Nothing Then
        clsStream(indicePlayer).CloseSource
        clsOutput(indicePlayer).StopPlay
    End If
    Set clsStream(indicePlayer) = Nothing
    
    'abrir segun formato XXXX
    'Set clsStream = streamFromExt(strFile)
    'If clsStream Is Nothing Then
    '    Debug.Print "Format nicht unterstützt"
    '    Exit Function
    'End If
    
    Set clsOutput(indicePlayer) = clsMonoLoader.CreateSoundOut()
    Set clsStream(indicePlayer) = New tbrDX8.StreamMP3
    
    'puedo saber si un archivo sirve!!!
    If clsStream(indicePlayer).OpenSource(sArchivo) <> STREAM_OK Then
        Player_AbrirMP3 = 1: sArchivo = ""
        Exit Function
    End If

    If Not clsOutput(indicePlayer).Initialize(clsStream(indicePlayer), Me) Then
        Player_AbrirMP3 = 2: sArchivo = ""
        Exit Function
    End If
    'todo ok
    isOpen(indicePlayer) = 1
    Player_AbrirMP3 = 0
    
End Function

Public Function Player_PlayMP3(indicePlayer As Long, Optional VolumenIni As Long = 100, Optional PanIni As Long = 0)
    '0: todo ok
    '1: falla

    clsOutput(indicePlayer).Volume = VolumenIni
    clsOutput(indicePlayer).Pan = PanIni

    If Not clsOutput(indicePlayer).Play() Then
        Player_PlayMP3 = 1
        Exit Function
    End If
    Player_PlayMP3 = 0
    
    'primero evento de comienzo y luego PLAYED
    RaiseEvent CancionEmpeiza(clsOutput(indicePlayer).MyID)
    
    'prender el reloj si corresponde
    If TmrPlay.Enabled = False Then TmrPlay.Enabled = True
End Function

Public Function Player_PausaMP3(indicePlayer As Long)
    clsOutput(indicePlayer).Pause
End Function

Public Function Player_StopMP3(indicePlayer As Long)
    clsOutput(indicePlayer).StopPlay
End Function


'****************************************
'eventos enviados
'****************************************
'termino una cancion, envio el ID para que sepa cual termino!
Private Sub IDSOutCallback_OnEndOfStream(sndout As SoundOut)
    RaiseEvent CancionTermina(sndout.MyID)
End Sub

Private Sub IDSOutCallback_OnStatusChanged(sndout As SoundOut, ByVal udeStat As tbrDX8.DS_PlayState)
    Dim NewState As String
    Select Case udeStat
        Case STAT_PAUSING: NewState = "Pause"
        Case STAT_PLAYING: NewState = "Play"
        Case STAT_STOPPED: NewState = "Stop"
    End Select
    RaiseEvent CancionCambiaEstado(Get_Player_IndexFromSID(sndout.MyID), sndout.MyID, NewState)
End Sub

Private Sub IDSOutCallback_Samples(sndout As SoundOut, intSamples() As Integer, ByVal datalength As Long, ByVal channels As Integer)
    '¿? sera uno por cada sample, o sea a cada rato ¿?
End Sub

'****************************************
'VUMetro
'****************************************
'una barra por canal
' L/R Channel Peaks
' http://www.activevb.de/tipps/vb6tipps/tipp0406.html
Public Sub DrawPeaks(indicePlayer As Long, picVisL As PictureBox, picVisR As PictureBox)

    Dim intSamples(FFT_SAMPLES - 1) As Integer

    If Get_Info_BitsPorSample(indicePlayer) = 16 Then
        clsOutput(indicePlayer).CaptureSamples VarPtr(intSamples(0)), FFT_SAMPLES * 2
    End If

    Dim i As Long, j As Long
    Dim maxL As Long, maxR As Long
    Dim rcBand As RECT
    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    Dim hBrush As Long

    Static LastL As Long
    Static LastR As Long

    ' mono?
    If Get_Info_Canales(indicePlayer) = 1 Then
        For i = LBound(intSamples) To UBound(intSamples)
            If Abs(CLng(intSamples(i))) > maxL Then
                maxL = Abs(CLng(intSamples(i)))
                maxR = maxL
            End If
        Next
    Else ' stereo!
        For i = LBound(intSamples) To UBound(intSamples)
            If i Mod 2 Then
                If Abs(CLng(intSamples(i))) > maxR Then
                    maxR = Abs(CLng(intSamples(i)))
                End If
            Else
                If Abs(CLng(intSamples(i))) > maxL Then
                    maxL = Abs(CLng(intSamples(i)))
                End If
            End If
        Next
    End If

    ' smoother value
    maxL = (LastL + maxL) / 2
    maxR = (LastR + maxR) / 2

    '*********************************************
    ' dibujar las barras
    '*********************************************
    intRed = 255
    intBlue = 0
    intGreen = 255 - (maxL / 32767 * 255)

    hBrush = CreateSolidBrush(RGB(intRed, intGreen, intBlue))
    
    picVisL.Cls
    ' barra izquierda
    With rcBand
        .Right = (picVisL.ScaleWidth / 15) - 4
        .Left = 4
        'va desde 0 (maximo) a -96
        'en cero debe ser 0
        .Top = 4 + ((Abs(Fix(dBFS(maxL))) / 100) * (picVisL.ScaleHeight - 8))
        .Bottom = (picVisL.ScaleHeight / 15) - 4
    End With
    FillRect picVisL.hdc, rcBand, hBrush
    'los decibeles serían:
    '
    DeleteObject hBrush

    intRed = 255
    intBlue = 0
    intGreen = 255 - (maxR / 32767 * 255)

    hBrush = CreateSolidBrush(RGB(intRed, intGreen, intBlue))

    picVisR.Cls
    ' barra derecha
    'los decibeles serían:
    'dBFS(maxR)
    With rcBand
        .Right = (picVisR.ScaleWidth / 15) - 4
        .Left = 4
        'va desde 0 (maximo) a -96
        'en cero debe ser 0
        .Top = 4 + ((Abs(Fix(dBFS(maxR))) / 100) * (picVisR.ScaleHeight - 8))
        .Bottom = (picVisR.ScaleHeight / 15) - 4
    End With
    FillRect picVisR.hdc, rcBand, hBrush

    DeleteObject hBrush
    
    LastL = maxL
    LastR = maxR
End Sub

Private Function dBFS(ByVal Amplitude As Long) As Double
    If Amplitude = 0 Then
        dBFS = -96
    Else
        dBFS = 20 * ((Log(Abs(Amplitude) / 32768)) / Log(10))
    End If
End Function

'****************************************
'menos importantes
'****************************************
Public Function TotalPlacas() As Long
    TotalPlacas = mTotalPlacas 'puede ser menos 1 si no se inicio nada
End Function

Public Function Get_Player_IndexFromSID(sID As Long) As Long
    Get_Player_IndexFromSID = -1 'valor si no se encuentra
    Dim j As Long
    For j = 0 To 3
        If sID = Get_Info_ID(j) Then
            Get_Player_IndexFromSID = j
            Exit For
        End If
    Next j
End Function


Public Function Get_Placa_IndexFromSID(sID As String) As Long
    Get_Placa_IndexFromSID = -1 'valor si no se encuentra
    Dim j As Long
    For j = 1 To mTotalPlacas
        If sID = PA(j).sID Then
            Get_Placa_IndexFromSID = j
            Exit For
        End If
    Next j
End Function

Public Function Get_Placa_Nombre_ByID(sID As String) As String
    'ver que indice tiene y sacarlo de alli
    Dim j As Long
    j = Get_Placa_IndexFromSID(sID)
    If j > 0 Then
        Get_Placa_Nombre_ByID = Get_Placa_Nombre_ByIndex(j)
    Else
        Get_Placa_Nombre_ByID = "" 'pide algo que no existe!
    End If
End Function

Public Function Get_Placa_Nombre_ByIndex(Index As Long) As String
    'ver que no pida huevadas
    If Index > mTotalPlacas Then
        Get_Placa_Nombre_ByIndex = ""
    Else
        Get_Placa_Nombre_ByIndex = PA(Index).Nombre
    End If
End Function

Public Function Get_Placa_Descripcion_ByID(sID As String) As String
    'ver que indice tiene y sacarlo de alli
    Dim j As Long
    j = Get_Placa_IndexFromSID(sID)
    If j > 0 Then
        Get_Placa_Descripcion_ByID = Get_Placa_Descripcion_ByIndex(j)
    Else
        Get_Placa_Descripcion_ByID = "" 'pide algo que no existe!
    End If
End Function

Public Function Get_Placa_Descripcion_ByIndex(Index As Long) As String
    'ver que no pida huevadas
    If Index > mTotalPlacas Then
        Get_Placa_Descripcion_ByIndex = ""
    Else
        Get_Placa_Descripcion_ByIndex = PA(Index).Descripcion
    End If
End Function

Public Function Get_Placa_ID(Index As Long) As String
    'ver que no pida huevadas
    If Index > mTotalPlacas Then
        Get_Placa_ID = ""
    Else
        Get_Placa_ID = PA(Index).sID
    End If
End Function

Public Function Get_Placa_Iniciada(Index As Long) As Long
    'me dic e si se inicio esta placa
    'ver que no pida huevadas
    If Index > mTotalPlacas Then
        Get_Placa_Iniciada = ""
    Else
        Get_Placa_Iniciada = PA(Index).Iniciada
    End If
End Function

Private Function SecToStrMinute(ByVal seconds As Long) As String
    Dim minutes     As Long
    minutes = seconds \ 60
    seconds = seconds Mod 60
    SecToStrMinute = Format(minutes, "00") & ":" & Format(seconds, "00")
End Function

Public Function Get_Info_SampleRate(indicePlayer As Long) As Long
    Get_Info_SampleRate = clsOutput(indicePlayer).StreamInfo.samplerate 'XXX revisar que haya algo abierto!
End Function

Public Function Get_Info_Canales(indicePlayer As Long) As Long
    Get_Info_Canales = clsOutput(indicePlayer).StreamInfo.channels
End Function

Public Function Get_Info_BitsPorSample(indicePlayer As Long) As Long
    Get_Info_BitsPorSample = CLng(clsOutput(indicePlayer).StreamInfo.bitspersample)
End Function

Public Function Get_Info_BitRate(indicePlayer As Long) As Long
    Get_Info_BitRate = clsOutput(indicePlayer).StreamInfo.Bitrate
End Function

Public Function Get_Info_DuracionMiliSegundos(indicePlayer As Long) As Long
    Get_Info_DuracionMiliSegundos = clsOutput(indicePlayer).StreamInfo.Duration
End Function

Public Function Get_Info_DuracionTexto(indicePlayer As Long) As String
    Get_Info_DuracionTexto = SecToStrMinute((clsOutput(indicePlayer).StreamInfo.Duration / 1000))
End Function

Public Function Get_Info_ID(indicePlayer As Long) As Long
    Get_Info_ID = clsOutput(indicePlayer).MyID
End Function

Public Function Get_Tags(indicePlayer As Long) As String()
    Dim TMP() As String, U As Long: U = 1
    ReDim TMP(0)
    'lista de tags del archivo
    For Each clsTag(indicePlayer) In clsStream(indicePlayer).Info.Tags
        With clsTag(indicePlayer)
            ReDim Preserve TMP(U)
            TMP(U) = .TagName + " = " + CStr(.TagValue)
            U = U + 1
        End With
    Next
    Get_Tags = TMP
    
End Function

Public Function Player_File() As String
    Player_File = sArchivo
End Function

Public Function Player_ClearMemory(indicePlayer As Long)
    isOpen(indicePlayer) = 0
    Set clsOutput(indicePlayer) = Nothing
    Set clsStream(indicePlayer) = Nothing
    Set clsTag(indicePlayer) = Nothing
End Function

Private Sub TmrPlay_Timer()
    'el evento que esta corriendo llega solo a los que esta corriendo ...
    Dim mSEG As Long, totmSEG As Long, U As Long
    For U = 0 To 3
        If isOpen(U) = 0 Then GoTo SIG
        'DrawPeaks
        totmSEG = clsOutput(U).StreamInfo.Duration
        mSEG = clsOutput(U).StreamInfo.position
        If clsOutput(U).Status = STAT_PLAYING Then
            DrawPeaks U, Form1.picVis(U * 2), Form1.picVis(U * 2 + 1)
            RaiseEvent CancionCorriendo( _
                Get_Player_IndexFromSID(clsOutput(U).MyID), _
                clsOutput(U).MyID, SecToStrMinute(mSEG / 1000), _
                mSEG, totmSEG - mSEG)
        End If
SIG:
    Next U
End Sub

Public Sub Player_Set_Volumen(indicePlayer As Long, NewVolumen As Long)
    clsOutput(indicePlayer).Volume = NewVolumen
End Sub

Public Sub Player_Set_Pan(indicePlayer As Long, NewPan As Long)
    clsOutput(indicePlayer).Pan = NewPan
End Sub
