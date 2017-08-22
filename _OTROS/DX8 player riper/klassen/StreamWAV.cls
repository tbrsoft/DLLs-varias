VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' WAV Decoder with ACM support

Implements ISoundStream
Implements ISoundInfo

' WAV file
Private hInput      As hFile
' DATA and FMT Chunk
Private cnkData     As CHUNKINFO
Private cnkInfo     As CHUNKINFO
' WAVEFORMATEX
Private btWFX()     As Byte
Private udtWFX      As MMWAVEFORMATEX

Private clsTags     As StreamTags

' ***********************************
' * constructor
' ***********************************

Private Sub Class_Initialize()
    Set clsTags = New StreamTags
End Sub

Private Sub Class_Terminate()
    ISoundStream_CloseSource
End Sub

' ***********************************
' * Stream properties
' ***********************************

Private Property Get ISoundInfo_Bitrate() As Long
    ISoundInfo_Bitrate = udtWFX.nAvgBytesPerSec * 8
End Property

Private Property Get ISoundInfo_BitsPerSample() As Integer
    ISoundInfo_BitsPerSample = udtWFX.wBitsPerSample
End Property

Private Property Get ISoundInfo_Channels() As Integer
    ISoundInfo_Channels = udtWFX.nChannels
End Property

Private Property Get ISoundInfo_Force16Bit() As Boolean
    ISoundInfo_Force16Bit = False
End Property

Private Property Get ISoundInfo_Samplerate() As Long
    ISoundInfo_Samplerate = udtWFX.nSamplesPerSec
End Property

Private Property Get ISoundInfo_Tags() As StreamTags
    Set ISoundInfo_Tags = clsTags
End Property

Private Property Get ISoundStream_EndOfStream() As Boolean
    ISoundStream_EndOfStream = FilePosition(hInput) >= (cnkData.Start + cnkData.Length)
End Property

Private Property Get ISoundStream_Extensions() As String()
    Dim strExt(0)   As String

    strExt(0) = "wav"
    ISoundStream_Extensions = strExt
End Property

Private Property Get ISoundStream_Info() As ISoundInfo
    Set ISoundStream_Info = Me
End Property

Private Property Get ISoundStream_Seekable() As Boolean
    ISoundStream_Seekable = True
End Property

Private Property Get ISoundStream_WFXPointer() As Long
    ISoundStream_WFXPointer = VarPtr(btWFX(0))
End Property

Private Property Get ISoundInfo_Duration() As Long
    ISoundInfo_Duration = (cnkData.Length) / udtWFX.nAvgBytesPerSec * 1000
End Property

Private Property Get ISoundInfo_Position() As Long
    ISoundInfo_Position = (FilePosition(hInput) - cnkData.Start) / udtWFX.nAvgBytesPerSec * 1000
End Property

' ***********************************
' * main functions
' ***********************************

Private Function ISoundStream_CloseSource() As STREAM_ERR
    FileClose hInput
End Function

Private Function ISoundStream_OpenSource( _
    ByVal Source As String _
) As STREAM_ERR

    If Not IsValidFile(Source) Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    cnkData = GetWavChunkPos(Source, "data")   ' 'data' Chunk
    cnkInfo = GetWavChunkPos(Source, "fmt ")   ' 'fmt ' Chunk

    ' valid Chunks?
    If cnkData.Start = 0 Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    If cnkInfo.Start = 0 Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    If cnkInfo.Length < 16 Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    hInput = FileOpen(Source, GENERIC_READ, FILE_SHARE_READ)
    If hInput.handle = INVALID_HANDLE Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    ' read Info Chunk
    ReDim btWFX(cnkInfo.Length - 1) As Byte
    FileSeek hInput, cnkInfo.Start, FILE_BEGIN
    FileRead hInput, VarPtr(btWFX(0)), cnkInfo.Length

    DXCopyMemory udtWFX, btWFX(0), Len(udtWFX)

    ' seek to the start of the data chunk
    FileSeek hInput, cnkData.Start, FILE_BEGIN

    ISoundStream_OpenSource = STREAM_OK
End Function

Private Function ISoundStream_Read( _
    ByVal pBuffer As Long, _
    ByVal cBytes As Long, _
    cBytesRead As Long _
) As STREAM_ERR

    ISoundStream_Read = STREAM_OK

    If hInput.handle = INVALID_HANDLE Then
        ISoundStream_Read = STREAM_INVALID_SOURCE
        Exit Function
    End If

    If FilePosition(hInput) > (cnkData.Start + cnkData.Length) Then
        ISoundStream_Read = STREAM_END_OF_STREAM
        cBytesRead = 0
        Exit Function
    End If

    If FilePosition(hInput) + cBytes > (cnkData.Start + cnkData.Length) Then
        cBytes = (cnkData.Start + cnkData.Length) - FilePosition(hInput)
        ISoundStream_Read = STREAM_END_OF_STREAM
    End If

    cBytesRead = FileRead(hInput, pBuffer, cBytes)
End Function

Private Function ISoundStream_SeekTo( _
    ByVal value As Long, _
    mode As SEEK_MODE _
) As STREAM_ERR

    Dim lngBytes    As Long

    If hInput.handle = INVALID_HANDLE Then Exit Function

    Select Case mode

        Case SEEK_PERCENT
            If value < 0 Or value > 100 Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            lngBytes = value / 100 * cnkData.Length

        Case SEEK_SECONDS
            If value < 0 Or value > (ISoundInfo_Duration / 1000) Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            lngBytes = udtWFX.nAvgBytesPerSec * value

    End Select

    If value = 0 Then
        FileSeek hInput, cnkData.Start, FILE_BEGIN
    Else
        lngBytes = AlignBytes(lngBytes) + cnkData.Start
        FileSeek hInput, lngBytes, FILE_BEGIN
    End If

    ISoundStream_SeekTo = STREAM_OK
End Function

' ***********************************
' * helpers
' ***********************************

Private Function AlignBytes( _
    ByVal bytes As Long _
) As Long

    AlignBytes = bytes - (bytes Mod udtWFX.nBlockAlign)
End Function
