VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' manager for ASPI, SPTI and IOCTLs
'
' this is the actual CDA stream,
' StreamCDA is just a wrapper

Implements ISoundStream
Implements ISoundInfo

Private Const samplerate        As Long = 44100
Private Const channels          As Long = 2
Private Const bitspersample     As Long = 16
Private Const BlockAlign        As Long = 4
Private Const BytesPerSecond    As Long = samplerate * BlockAlign

Private Const SectorSize        As Long = 2352
Private Const SectorsPerRead    As Long = 4

Private clsInterfaces()         As ICDAccess
Private clsCD                   As ICDAccess
Private lngInterfaceCount       As Long

Private btTrack                 As Byte
Private clsTrack                As CDTOCTrack
Private lngDuration             As Long
Private lngPosition             As Long

Private btBuffer(SectorSize * SectorsPerRead - 1) As Byte
Private lngBufPos               As Long
Private lngBufLen               As Long

Private udtWFX                  As MMWAVEFORMATEX

Private Function FillBuffer() As Boolean
    If ISoundStream_EndOfStream Then
        Exit Function
    End If

    If clsCD.ReadSectorsRAW(clsTrack.StartLBA + lngPosition, _
                            SectorsPerRead, _
                            VarPtr(btBuffer(0))) <> STREAM_OK Then

        Exit Function
    End If

    lngPosition = lngPosition + SectorsPerRead

    lngBufPos = 0
    lngBufLen = SectorSize * SectorsPerRead

    FillBuffer = True
End Function

Public Function SelectTrack( _
    ByVal track As Byte _
) As STREAM_ERR

    If track < 1 Or track > GetTOC.TrackCount - 1 Then
        SelectTrack = STREAM_OUT_OF_RANGE
        Exit Function
    End If

    Set clsTrack = GetTOC.track(track)
    lngDuration = GetTOC.track(track + 1).StartLBA - clsTrack.StartLBA
    lngPosition = 0

    SelectTrack = STREAM_OK
End Function

Public Function GetTOC() As CDTOC
    Dim clsTOC  As CDTOC

    If STREAM_OK = clsCD.GetTOC(clsTOC) Then
        Set GetTOC = clsTOC
    End If
End Function

Public Function IsDiscPresent() As Boolean
    If clsCD Is Nothing Then Exit Function
    IsDiscPresent = clsCD.IsDiscPresent
End Function

Public Function CloseCD() As Boolean
    If clsCD Is Nothing Then Exit Function
    CloseCD = clsCD.CloseCD = STREAM_OK
End Function

Public Function OpenCD() As Boolean
    If clsCD Is Nothing Then Exit Function
    OpenCD = clsCD.OpenCD = STREAM_OK
End Function

Public Property Get DeviceCount() As Long
    If clsCD Is Nothing Then Exit Property
    DeviceCount = clsCD.DeviceCount
End Property

Public Property Get DeviceName( _
    ByVal index As Long _
) As String

    If clsCD Is Nothing Then Exit Property
    DeviceName = clsCD.DeviceName(index)
End Property

Public Property Get DeviceChar( _
    ByVal index As Long _
) As String

    If clsCD Is Nothing Then Exit Property
    DeviceChar = clsCD.DeviceChar(index)
End Property

Public Function SelectDevice( _
    ByVal index As Long _
) As Boolean

    If clsCD Is Nothing Then Exit Function
    SelectDevice = clsCD.SelectDevice(index) = STREAM_OK
End Function

Public Function InitManager( _
    Optional DisableIOCTL As Boolean _
) As Boolean

    Dim i   As Long

    ' search for a working interface

    lngInterfaceCount = 0
    AddInterface New CDAccessSPTI
    If Not DisableIOCTL Then
        AddInterface New CDAccessIOCTL
    End If
    AddInterface New CDAccessASPI

    For i = 0 To lngInterfaceCount - 1
        Set clsCD = clsInterfaces(i)

        If clsCD.InitMe() = STREAM_OK Then
            Exit For
        Else
            Set clsCD = Nothing
            Set clsInterfaces(i) = Nothing
        End If
    Next

    If i = lngInterfaceCount Then
        Set clsCD = Nothing
        Exit Function
    End If

    InitManager = True
End Function

Private Sub AddInterface( _
    c As ICDAccess _
)

    ReDim Preserve clsInterfaces(lngInterfaceCount) As ICDAccess
    Set clsInterfaces(lngInterfaceCount) = c
    lngInterfaceCount = lngInterfaceCount + 1
End Sub

Private Sub Class_Initialize()
    modDevIoCtl.InitIOCTLs
    udtWFX = CreateWFX(samplerate, channels, bitspersample)
End Sub

Private Sub Class_Terminate()
    Dim i   As Integer

    clsCD.DeInitMe
    Set clsCD = Nothing

'    For i = 0 To lngInterfaceCount - 1
'        Set clsInterfaces(i) = Nothing
'    Next
End Sub

Private Property Get ISoundInfo_Force16Bit() As Boolean
    ISoundInfo_Force16Bit = True
End Property

Private Property Get ISoundInfo_Tags() As StreamTags
    '
End Property

Private Property Get ISoundStream_EndOfStream() As Boolean
    ISoundStream_EndOfStream = clsTrack.StartLBA + lngPosition >= clsTrack.StartLBA + lngDuration
End Property

Private Property Get ISoundStream_Extensions() As String()
    Dim strExt(0)   As String

    strExt(0) = "cda"
    ISoundStream_Extensions = strExt
End Property

Private Function ISoundStream_CloseSource() As STREAM_ERR
    ISoundStream_CloseSource = STREAM_OK
End Function

Private Property Get ISoundStream_Info() As ISoundInfo
    Set ISoundStream_Info = Me
End Property

Private Function ISoundStream_OpenSource( _
    ByVal Source As String _
) As STREAM_ERR

    ' CDManager does the opening itself
    ISoundStream_OpenSource = STREAM_INVALID_SOURCE
End Function

Private Function ISoundStream_Read( _
    ByVal pointer As Long, _
    ByVal cBytes As Long, _
    cBytesRead As Long _
) As STREAM_ERR

    Do While cBytesRead < cBytes

        If lngBufLen - lngBufPos < cBytes - cBytesRead Then

            If 0 < (lngBufLen - lngBufPos) Then
                DXCopyMemory ByVal pointer + cBytesRead, btBuffer(lngBufPos), lngBufLen - lngBufPos
                cBytesRead = cBytesRead + (lngBufLen - lngBufPos)
            End If

            If Not FillBuffer Then
                ISoundStream_Read = STREAM_END_OF_STREAM
                Exit Function
            End If

        Else

            DXCopyMemory ByVal pointer + cBytesRead, btBuffer(lngBufPos), cBytes - cBytesRead
            lngBufPos = lngBufPos + (cBytes - cBytesRead)
            cBytesRead = cBytesRead + (cBytes - cBytesRead)

        End If

    Loop

    ISoundStream_Read = STREAM_OK
End Function

Private Property Get ISoundStream_Seekable() As Boolean
    ISoundStream_Seekable = True
End Property

Private Function ISoundStream_SeekTo( _
    ByVal value As Long, _
    mode As SEEK_MODE _
) As STREAM_ERR

    ZeroMemory btBuffer(0), UBound(btBuffer) + 1

    Select Case mode

        Case SEEK_PERCENT
            If value < 0 Or value > 100 Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            lngPosition = value * (ISoundInfo_Duration / 1000) / 100

        Case SEEK_SECONDS
            If value < 0 Or value > (ISoundInfo_Duration / 1000) Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            lngPosition = (value * BytesPerSecond) / SectorSize

    End Select

    ISoundStream_SeekTo = STREAM_OK
End Function

Private Property Get ISoundStream_StreamInfo() As ISoundInfo
    Set ISoundStream_StreamInfo = Me
End Property

Private Property Get ISoundInfo_Bitrate() As Long
    ISoundInfo_Bitrate = BytesPerSecond * 8&
End Property

Private Property Get ISoundInfo_BitsPerSample() As Integer
    ISoundInfo_BitsPerSample = bitspersample
End Property

Private Property Get ISoundInfo_Channels() As Integer
    ISoundInfo_Channels = channels
End Property

Private Property Get ISoundInfo_Duration() As Long
    ISoundInfo_Duration = ((lngDuration * SectorSize) / BytesPerSecond) * 1000
End Property

Private Property Get ISoundInfo_Position() As Long
    ISoundInfo_Position = ((lngPosition * SectorSize) / BytesPerSecond) * 1000
End Property

Private Property Get ISoundInfo_Samplerate() As Long
    ISoundInfo_Samplerate = samplerate
End Property

Private Property Get ISoundInfo_Seekable() As Boolean
    ISoundInfo_Seekable = True
End Property

Private Property Get ISoundInfo_TagsExist() As Boolean
    ISoundInfo_TagsExist = False
End Property

Private Property Get ISoundInfo_Title() As String
    '
End Property

Private Property Get ISoundInfo_Album() As String
    '
End Property

Private Property Get ISoundInfo_Artist() As String
    '
End Property

Private Property Get ISoundStream_WFXPointer() As Long
    ISoundStream_WFXPointer = VarPtr(udtWFX)
End Property
