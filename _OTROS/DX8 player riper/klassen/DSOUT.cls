VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' DirectSound Stream Output

Implements IDXCallback
Implements IDSPCallback

Private Declare Sub ZeroMem Lib "kernel32" _
Alias "RtlZeroMemory" ( _
    pData As Any, _
    ByVal dwLen As Long _
)

Private Declare Sub FillMem Lib "kernel32.dll" _
Alias "RtlFillMemory" ( _
    pData As Any, _
    ByVal dwLen As Long, _
    ByVal Fill As Byte _
)

' The DirectSound buffer is BUFFER_LENGTH_MS
' ms big. Because we stream, it get's segmented
' into BUFFER_SEGMENTS segments.
' Every segment is BUFFER_SEGMENT_MS ms big.
'
' 2 segments seem to be the best solution,
' for example I get problems if I hold the mouse
' button on the titlebar, the segment order
' seems to get shuffled then...
Private Const BUFFER_LENGTH_MS  As Long = 2000
Private Const BUFFER_SEGMENTS   As Long = 2
Private Const BUFFER_SEGMENT_MS As Long = BUFFER_LENGTH_MS / BUFFER_SEGMENTS

Public Enum DS_PlayState
    STAT_STOPPED
    STAT_PLAYING
    STAT_PAUSING
End Enum

Public Enum DS_Unit
    UNIT_DECIBEL
    UNIT_LINEAR
End Enum

Private Const fLog10            As Single = 0.434294481903252

' DirectSound Buffer
Private DXBuffer                As DirectSoundSecondaryBuffer8
Private DirectSound             As DirectSound8

' Segment Callback
Private frmTimeCB               As frmDXCallback
' Info Callback
Private clsInfoCB               As IDSOutCallback
' Data stream
Private clsData                 As PCMPreparator
Private clsStream               As ISoundStream

' current format, used by DXBuffer
Private udtFormat               As DSBUFFERDESC
' get notified when these positions
' are reached in the buffer
Private udtNotify()             As DSBPOSITIONNOTIFY

' segment size in bytes
Private lngSegmentSize          As Long
' buffer size in bytes
Private lngBufferSize           As Long

' write position in bytes
Private lngBufferPos            As Long
' notify events
Private lngNotifyEvent          As Long
Private lngNotifyEventUBound    As Long

' last filled segment
Private lngLastSegment          As Long

' DirectSound status and EndOfStream flag
Private udeStatus               As DS_PlayState
Private udeVolUnit              As DS_Unit

' elapes milliseconds
Private lngCurrentMS            As Long

' copy of the DirectSound databuffer
Private btBuffer()              As Byte

Private blnEOS                  As Boolean

' DSP only for 16 Bit samples,
' so convert 8 Bit streams to 16 bit
Private blnForce16Bit           As Boolean

Private lngVolume               As Long
Private lngPan                  As Long

Friend Property Set DS( _
    clsDS As DirectSound8 _
)

    Set DirectSound = clsDS
End Property

Public Property Get channels() As Integer
    channels = clsData.OutputChannels
End Property

Public Property Get samplerate() As Long
    samplerate = clsData.OutputSamplerate
End Property

Public Property Get bitspersample() As Integer
    bitspersample = clsData.OutputBitsPerSample
End Property

Public Function CaptureSamples( _
    ByVal pBuffer As Long, _
    ByVal bytes As Long _
) As Boolean

    Dim udtPos      As DSCURSORS
    Dim lngPNext    As Long
    Dim lngCapLen   As Long
    Dim lngCapStart As Long

    If Status <> STAT_PLAYING Then
        Exit Function
    End If

    If bytes > MSToBytes(BUFFER_LENGTH_MS) Then
        Exit Function
    End If

    DXBuffer.GetCurrentPosition udtPos

    lngCapStart = udtPos.lPlay

    If lngCapStart + bytes > MSToBytes(BUFFER_LENGTH_MS) Then
        ' not enough data till the end of the buffer,
        ' get some more data from the beginning of the buffer
        DXCopyMemory ByVal pBuffer, btBuffer(lngCapStart), MSToBytes(BUFFER_LENGTH_MS) - lngCapStart
        lngPNext = MSToBytes(BUFFER_LENGTH_MS) - lngCapStart
        bytes = bytes - (MSToBytes(BUFFER_LENGTH_MS) - lngCapStart)
        DXCopyMemory ByVal pBuffer + lngPNext, btBuffer(0), bytes
    Else
        DXCopyMemory ByVal pBuffer, btBuffer(lngCapStart), bytes
    End If

    CaptureSamples = True
End Function

Public Function SeekTo( _
    ByVal value As Integer, _
    mode As SEEK_MODE _
) As Boolean

    If clsStream.SeekTo(value, mode) <> STREAM_OK Then
        Exit Function
    End If

    ' reset the buffer
    DXBuffer.Stop
    lngLastSegment = 0
    lngBufferPos = 0
    DXBuffer.SetCurrentPosition 0

    clsData.ResetACM

    ' if we cut the play, continue
    If Status = STAT_PLAYING Then
        GetAndWriteSegment
        DXBuffer.Play DSBPLAY_LOOPING
    End If

    lngCurrentMS = clsStream.Info.position - BUFFER_LENGTH_MS
    If lngCurrentMS < 0 Then lngCurrentMS = 0

    SeekTo = True
End Function

Public Property Get position() As Long
    Dim udtPos  As DSCURSORS
    Dim lngTime As Long
    Static lastx As Long

    DXBuffer.GetCurrentPosition udtPos

    udtPos.lPlay = BytesToMS(udtPos.lPlay)

    position = lngCurrentMS + udtPos.lPlay
End Property

Public Property Get Duration() As Long
    Dim lngDuration As Long

    lngDuration = clsStream.Info.Duration
    lngDuration = lngDuration + (lngDuration Mod BUFFER_SEGMENT_MS)

    Duration = lngDuration
End Property

Public Property Get Force16Bit() As Boolean
    Force16Bit = blnForce16Bit
End Property

Public Property Let Force16Bit( _
    ByVal bln As Boolean _
)

    blnForce16Bit = bln
End Property

Public Property Get StreamInfo() As ISoundInfo
    Set StreamInfo = clsStream
End Property

Public Property Get Status() As DS_PlayState
    Status = udeStatus
End Property

Private Property Let Status( _
    ude As DS_PlayState _
)

    udeStatus = ude
    clsInfoCB.OnStatusChanged Me, udeStatus
End Property

' volume unit (dB, linear)
Public Property Let VolumeUnit( _
    ude As DS_Unit _
)

    udeVolUnit = ude
End Property

Public Property Get VolumeUnit() As DS_Unit
    VolumeUnit = udeVolUnit
End Property

' unique ID of the buffer
Public Property Get MyID() As Long
    MyID = ObjPtr(Me)
End Property

Public Property Get Pan() As Long
    Pan = DXBuffer.GetPan()
End Property

' relative volume of the channels
' in -10000 to +10000 100-dB
'
'     0:  (0 dB) both channels at full volume
' -5000: (50 dB) left channel at full vol,
'                right channel at -50 dB
' +2700: (27 dB) left channel at -27 dB,
'                right channel at full vol
'
Public Property Let Pan(ByVal lngValue As Long)

    If lngValue < DSBPAN_LEFT Then lngValue = DSBPAN_LEFT
    If lngValue > DSBPAN_RIGHT Then lngValue = DSBPAN_RIGHT

    lngPan = lngValue

    If Not DXBuffer Is Nothing Then DXBuffer.SetPan lngPan
    
End Property

Public Property Get Volume() As Long
    Dim lngVol  As Long

    lngVol = DXBuffer.GetVolume()

    Select Case udeVolUnit
        Case UNIT_DECIBEL
            ' 100-dB to dB
            Volume = 100 + (lngVol / 100)
        Case UNIT_LINEAR
            ' 100-dB to linear unit
            If lngVol = DSBVOLUME_MIN Then
                ' here we'd get 1, not beautiful
                Volume = 0
            Else
                Volume = 10 ^ (lngVol / 50 / 100) * 100
            End If
    End Select
End Property

Public Property Let Volume(ByVal lngVol As Long)

    Select Case udeVolUnit
        Case UNIT_DECIBEL
            ' dB to 100-dB
            lngVol = -(100 - lngVol) * 100
        Case UNIT_LINEAR
            ' linear unit to 100-dB
            If lngVol <= 0 Then lngVol = 1
            lngVol = 100 * 50 * Log10(lngVol / 100)
    End Select

    If lngVol < DSBVOLUME_MIN Then _
        lngVol = DSBVOLUME_MIN
    If lngVol > DSBVOLUME_MAX Then
        lngVol = DSBVOLUME_MAX
    End If

    lngVolume = lngVol

    If Not DXBuffer Is Nothing Then
        DXBuffer.SetVolume lngVol
    End If
End Property

Public Function Play() As Boolean
    On Error GoTo ErrorHandler

    If clsStream Is Nothing Then
        Exit Function
    End If

    Select Case Status
        Case STAT_PAUSING
            Pause
        Case STAT_PLAYING
            '
        Case STAT_STOPPED
            GetAndWriteSegment
            DXBuffer.Play DSBPLAY_LOOPING
            Status = STAT_PLAYING
    End Select

    Play = True

ErrorHandler:
End Function

Public Sub Pause()
    If clsStream Is Nothing Then
        Exit Sub
    End If

    Select Case Status
        Case STAT_PLAYING
            DXBuffer.Stop
            Status = STAT_PAUSING
        Case STAT_PAUSING
            DXBuffer.Play DSBPLAY_LOOPING
            Status = STAT_PLAYING
    End Select
End Sub

Public Sub StopPlay()
    If clsStream Is Nothing Then
        Exit Sub
    End If

    If Status = STAT_STOPPED Then
        Exit Sub
    End If

    DXBuffer.Stop
    DXBuffer.SetCurrentPosition 0

    lngBufferPos = 0
    lngLastSegment = 0
    lngCurrentMS = 0

    clsData.ResetACM

    Status = STAT_STOPPED
End Sub

Public Function Initialize(ByVal stream As ISoundStream, ByRef CallBack As IDSOutCallback) As Boolean

    On Error GoTo ErrorHandler

    Dim i   As Long

    ' clean up old stream, if opened
    If Not clsStream Is Nothing Then
        clsData.CloseConverter
    End If

    Set clsStream = stream

    ' init ACM bridge
    If Not clsData.InitConversion(clsStream, Me, blnForce16Bit) Then
        Exit Function
    End If

    With udtFormat
        With .fxFormat
            .nSize = LenB(udtFormat.fxFormat)
            .nFormatTag = WAVE_FORMAT_PCM
            .nChannels = clsData.OutputChannels
            .lSamplesPerSec = clsData.OutputSamplerate
            .nBitsPerSample = clsData.OutputBitsPerSample

            .nBlockAlign = .nChannels * (.nBitsPerSample / 8)
            .lAvgBytesPerSec = .nBlockAlign * .lSamplesPerSec
        End With

        .lBufferBytes = MSToBytes(BUFFER_LENGTH_MS)

        .lFlags = DSBCAPS_CTRLVOLUME Or _
                  DSBCAPS_CTRLPAN Or _
                  DSBCAPS_CTRLPOSITIONNOTIFY Or _
                  DSBCAPS_GETCURRENTPOSITION2 Or _
                  DSBCAPS_GLOBALFOCUS
    End With

    Set DXBuffer = DirectSound.CreateSoundBuffer(udtFormat)

    lngBufferSize = udtFormat.lBufferBytes
    lngSegmentSize = lngBufferSize / BUFFER_SEGMENTS

    ReDim btBuffer(lngBufferSize - 1) As Byte

    ' event for position notification
    lngNotifyEvent = frmDXCallback.CreateEvent
    ' event for increasing ms counter
    lngNotifyEventUBound = frmDXCallback.CreateEvent

    ' each segment center is a notify position
    ReDim udtNotify(BUFFER_SEGMENTS) As DSBPOSITIONNOTIFY

    For i = 0 To BUFFER_SEGMENTS - 1
        With udtNotify(i)
            .lOffset = MSToBytes(SegmentCenter(i + 1))
            .hEventNotify = lngNotifyEvent
        End With
    Next

    ' notification 5 ms before buffer ubound
    With udtNotify(i)
        .lOffset = MSToBytes(SegmentUBound(BUFFER_SEGMENTS) - 5)
        .hEventNotify = lngNotifyEventUBound
    End With

    DXBuffer.SetNotificationPositions BUFFER_SEGMENTS + 1, udtNotify

    Set clsInfoCB = CallBack

    lngCurrentMS = 0
    lngBufferPos = 0
    lngLastSegment = 0

    Volume = lngVolume
    Pan = lngPan

    Initialize = True

ErrorHandler:
End Function

Private Sub Class_Initialize()
    Set frmTimeCB = New frmDXCallback
    Set clsData = New PCMPreparator

    modEventManager.AddEvent Me

    ' events don't get fired in the constructor
    udeStatus = STAT_STOPPED
    udeVolUnit = UNIT_LINEAR
End Sub

' write data to the DX buffer
Private Sub WriteToBuffer( _
    ByVal p As Long, _
    ByVal Length As Long _
)

    DXBuffer.WriteBuffer lngBufferPos, _
                         Length, _
                         ByVal p, _
                         DSBLOCK_DEFAULT

    lngBufferPos = lngBufferPos + Length
    If lngBufferPos >= lngBufferSize Then
        lngBufferPos = 0
    End If
End Sub

Private Sub GetAndWriteSegment()
    Dim lngWritten  As Long

    ' end of stream, cut it!
    If blnEOS Then
        DXBuffer.Stop
        DXBuffer.SetCurrentPosition 0
        lngLastSegment = 0
        Status = STAT_STOPPED
        clsData.ResetACM
        blnEOS = False
        'LooseCodePath1P clsInfoCB, "OnEndOfStream", Me
        clsInfoCB.OnEndOfStream Me
        Exit Sub
    End If

    ' read a segment from the stream
     clsData.GetSamples VarPtr(btBuffer(lngBufferPos)), _
                        lngSegmentSize, _
                        lngWritten

    ' if we havn't got enough data to fill the
    ' complete segment, fill the rest with silence
    If lngWritten < lngSegmentSize Then
        If udtFormat.fxFormat.nBitsPerSample = 16 Then
            FillMem btBuffer(lngBufferPos + lngWritten), lngSegmentSize - lngWritten, 0
        Else
            FillMem btBuffer(lngBufferPos + lngWritten), lngSegmentSize - lngWritten, 128
        End If
    End If

    ' write read data to the DirectSound buffer
    WriteToBuffer VarPtr(btBuffer(lngBufferPos)), _
                  lngSegmentSize

    ' nothing written = EOS
    blnEOS = lngWritten = 0
End Sub

Private Sub Class_Terminate()
    ' close connection to the event manager
    modEventManager.RemEvent lngNotifyEvent
    modEventManager.RemEvent lngNotifyEventUBound

    clsData.CloseConverter
End Sub

' redirect samples
Private Sub IDSPCallback_Samples( _
    intSamples() As Integer, _
    ByVal datalength As Long, _
    ByVal channels As Integer _
)

    clsInfoCB.Samples Me, intSamples, _
                      datalength, channels
End Sub

' DX Events created by this class
Private Property Get IDXCallback_MyEvents() As Long()
    Dim lngEvents(1)    As Long

    lngEvents(0) = lngNotifyEvent
    lngEvents(1) = lngNotifyEventUBound

    IDXCallback_MyEvents = lngEvents
End Property

' DX Event raised
Private Sub IDXCallback_OnEvent( _
    ByVal eventid As Long _
)

    Dim udtPos  As DSCURSORS
    Dim i       As Long

    If Status = STAT_STOPPED Then Exit Sub

    DXBuffer.GetCurrentPosition udtPos

    ' increase the ms counter at each upper
    ' segment bound
    If eventid = lngNotifyEventUBound Then
        lngCurrentMS = lngCurrentMS + BUFFER_LENGTH_MS
        Exit Sub
    End If

    ' Pause/Resume seems to confuse Event callbacks.
    ' An event can be fired immediately after
    ' resuming playback. However it gets fired
    ' again when reaching the center of the segment.
    ' Consequence is a time-displaced segment.
    For i = 1 To BUFFER_SEGMENTS
        ' notification is in a valid timeframe?
        If udtPos.lPlay > MSToBytes(SegmentCenter(i)) Then
            If udtPos.lPlay < MSToBytes(SegmentLBound(i + 1)) Then
                ' do not fill the same segment
                ' a second time
                If lngLastSegment = i Then Exit Sub
                lngLastSegment = i
                GetAndWriteSegment
                Exit Sub
            End If
        End If
    Next
End Sub

Private Function BytesToMS( _
    ByVal bytes As Long _
) As Long

    BytesToMS = bytes / udtFormat.fxFormat.lAvgBytesPerSec * 1000
End Function

Private Function MSToBytes( _
    ByVal ms As Long _
) As Long

    MSToBytes = ms / 1000 * udtFormat.fxFormat.lAvgBytesPerSec
End Function

Private Function SegmentLBound( _
    ByVal i As Long _
) As Long

    SegmentLBound = SegmentUBound(i - 1)
End Function

Private Function SegmentUBound( _
    ByVal i As Long _
) As Long

    SegmentUBound = (BUFFER_LENGTH_MS / BUFFER_SEGMENTS) * i
End Function

Private Function SegmentCenter( _
    ByVal i As Long _
) As Long

    SegmentCenter = SegmentUBound(i) - (BUFFER_LENGTH_MS / BUFFER_SEGMENTS / 2)
End Function

Private Function Log10( _
    ByVal value As Single _
) As Single

    Log10 = Log(value) * fLog10
End Function
