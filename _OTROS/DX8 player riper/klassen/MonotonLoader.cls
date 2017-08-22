VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MonotonLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private blnDXReady  As Boolean
Private clsDSPrim   As DirectSoundPrimaryBuffer8
Private clsDS       As DirectSound8
Private clsDSEnum   As DirectSoundEnum8

Public Property Get DeviceGUID( _
    index As Long _
) As String

    DeviceGUID = clsDSEnum.GetGuid(index)
End Property

Public Property Get DeviceName( _
    index As Long _
) As String

    DeviceName = clsDSEnum.GetName(index)
End Property

Public Property Get DeviceDescription( _
    index As Long _
) As String

    DeviceDescription = clsDSEnum.GetDescription(index)
End Property

Public Property Get DeviceCount() As Long
    DeviceCount = clsDSEnum.GetCount
End Property

Public Function CreateSoundOut() As SoundOut
    Dim dsout   As SoundOut
    Set dsout = New SoundOut
    Set dsout.DS = clsDS
    Set CreateSoundOut = dsout
End Function

Public Function Initialize( _
    ByVal samplerate As Long, _
    ByVal channels As Integer, _
    ByVal bitspersample As Integer, _
    Optional ByVal devguid As String = "" _
) As Boolean

    On Error GoTo ErrorHandler

    If Not blnDXReady Then Exit Function

    Set clsDSPrim = Nothing
    Set clsDS = Nothing

    Set clsDS = DirectX.DirectSoundCreate(devguid)

    clsDS.SetCooperativeLevel frmDXCallback.hWnd, _
                              DSSCL_PRIORITY

    Set clsDSPrim = CreatePrimBuf(samplerate, _
                                  channels, _
                                  bitspersample, _
                                  clsDS)

    Initialize = True

ErrorHandler:
End Function

Private Sub Class_Initialize()
    blnDXReady = modPrimaryBuffer.InitializeDirectX
    modEventManager.InitEventManager

    If blnDXReady Then
        Set clsDSEnum = modPrimaryBuffer.DirectX.GetDSEnum
    End If
End Sub

Private Function CreatePrimBuf( _
    ByVal samplerate As Long, _
    ByVal channels As Integer, _
    ByVal bitspersample As Long, _
    ByRef DirectSound As DirectSound8 _
) As DirectSoundPrimaryBuffer8

    Dim udtFormat       As DSBUFFERDESC
    Dim DXPrimaryBuffer As DirectSoundPrimaryBuffer8

    With udtFormat
        With .fxFormat
            .lSamplesPerSec = samplerate
            .nChannels = channels
            .nBitsPerSample = bitspersample
            .nBlockAlign = .nChannels * (.nBitsPerSample / 8)
            .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
            .nFormatTag = WAVE_FORMAT_PCM
            .nSize = LenB(udtFormat.fxFormat)
        End With
        .lBufferBytes = 0
        .lFlags = DSBCAPS_PRIMARYBUFFER Or DSBCAPS_CTRLVOLUME
    End With

    Set DXPrimaryBuffer = DirectSound.CreatePrimarySoundBuffer(udtFormat)
    Set CreatePrimBuf = DXPrimaryBuffer
End Function
