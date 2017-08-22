VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' MPEG Layer 3 Decoder (Lame ACM Codec)

Implements ISoundStream
Implements ISoundInfo

Private Type XingHeader
    valid                   As Boolean
    Flags                   As Long
    Frames                  As Long
    bytes                   As Long
    toc(99)                 As Byte
    Quality                 As Long
End Type

Private Type FrameHeader
    sync(3)                 As Byte
    MPEGID                  As Byte
    layer                   As Byte
    ProtectBit              As Byte
    Bitrate                 As Byte
    samplerate              As Byte
    PaddingBit              As Byte
    PrivateBit              As Byte
    ChannelMode             As Byte
    ModeExt                 As Byte
    CpyrightBit             As Byte
    OriginalBit             As Byte
    emphasis                As Byte
End Type

Private Type MPEGLAYER3WAVEFORMAT
    wFormatTag              As Integer
    nChannels               As Integer
    nSamplesPerSec          As Long
    nAvgBytesPerSec         As Long
    nBlockAlign             As Integer
    wBitsPerSample          As Integer
    cbSize                  As Integer
    wID                     As Integer
    fdwFlags                As Long
    nBlockSize              As Integer
    nFramesPerBlock         As Integer
    nCodecDelay             As Integer
End Type

Private Const MPEGLAYER3_ID_UNKNOWN = 0
Private Const MPEGLAYER3_ID_MPEG = 1
Private Const MPEGLAYER3_ID_CONSTANTFRAMESIZE = 2

Private Const MPEGLAYER3_FLAG_PADDING_ISO = &H0
Private Const MPEGLAYER3_FLAG_PADDING_ON = &H1
Private Const MPEGLAYER3_FLAG_PADDING_OFF = &H2

Private Const WAVE_FORMAT_MPEGLAYER3 = &H55
Private Const MPEGLAYER3_WFX_EXTRA_BYTES = 12

Private Const MPEGLAYER3_BLOCK_SIZE = 522

Private Const FRAME_SYNC_REGION_KB = 64

Private Enum eXingFlags
    VBR_FRAMES = &H1                ' MP3 Frames
    VBR_BYTES = &H2                 ' MP3 Bytes
    VBR_TOC = &H4                   ' Table Of Contents
    VBR_SCALE = &H8                 ' MP3 quality
End Enum

Private Enum eMPEGID
    [MPEG 2.5] = &H0
    [MPEG 2] = &H2
    [MPEG 1] = &H3
End Enum

' MP3 file
Private hInput              As hFile
' offset of the first frame header
Private lngSyncPos          As Long

' Frame Header
Private udtFrameHdr         As FrameHeader
' XING Header (VBR)
Private udtXingHdr          As XingHeader

' ACM MP3 WFX
Private udtWFX              As MPEGLAYER3WAVEFORMAT

Private clsTags             As StreamTags

' *********************************
' * constructor
' *********************************

Private Sub Class_Initialize()
    Set clsTags = New StreamTags
End Sub

Private Sub Class_Terminate()
    ISoundStream_CloseSource
End Sub

' *********************************
' * Stream properties
' *********************************

Private Property Get ISoundInfo_Bitrate() As Long
    ISoundInfo_Bitrate = Bitrate * 1000
End Property

Private Property Get ISoundInfo_BitsPerSample() As Integer
    ISoundInfo_BitsPerSample = 16
End Property

Private Property Get ISoundInfo_Channels() As Integer
    ISoundInfo_Channels = channels
End Property

Private Property Get ISoundInfo_Duration() As Long
    ISoundInfo_Duration = Duration
End Property

Private Property Get ISoundInfo_Force16Bit() As Boolean
    ISoundInfo_Force16Bit = True
End Property

Private Property Get ISoundInfo_Position() As Long
    ISoundInfo_Position = CurrentTime
End Property

Private Property Get ISoundInfo_Samplerate() As Long
    ISoundInfo_Samplerate = samplerate
End Property

Private Property Get ISoundInfo_Tags() As StreamTags
    Set ISoundInfo_Tags = clsTags
End Property

Private Property Get ISoundStream_EndOfStream() As Boolean
    ISoundStream_EndOfStream = FileEnd(hInput)
End Property

Private Property Get ISoundStream_Extensions() As String()
    Dim strExt(0)   As String

    strExt(0) = "mp3"
    ISoundStream_Extensions = strExt
End Property

Private Property Get ISoundStream_Info() As ISoundInfo
    Set ISoundStream_Info = Me
End Property

Private Property Get ISoundStream_Seekable() As Boolean
    ISoundStream_Seekable = True
End Property

Private Property Get ISoundStream_WFXPointer() As Long
    ISoundStream_WFXPointer = VarPtr(udtWFX)
End Property

' *********************************
' * main functions
' *********************************

Private Function ISoundStream_CloseSource() As STREAM_ERR
    FileClose hInput
    ISoundStream_CloseSource = STREAM_OK
End Function

Private Function ISoundStream_OpenSource( _
    ByVal Source As String _
) As STREAM_ERR

    Dim sync(3) As Byte
    Dim id3v2   As ID3Tag
    Dim i       As Long

    If Not IsValidFile(Source) Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    If hInput.handle <> INVALID_HANDLE Then
        ISoundStream_CloseSource
    End If

    GetTags Source
    modID3Tags.ReadID3v2 Source, id3v2

    hInput = FileOpen(Source, GENERIC_READ, FILE_SHARE_READ)
    If hInput.handle = INVALID_HANDLE Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    ' skip ID3v2, if at beginning of the file
    If id3v2.position < FRAME_SYNC_REGION_KB * 1024& And id3v2.size > 1 Then
        lngSyncPos = SearchFirstSync(FRAME_SYNC_REGION_KB, True, id3v2.position + id3v2.size)
    Else
        lngSyncPos = SearchFirstSync(FRAME_SYNC_REGION_KB, True, 0)
    End If

    If lngSyncPos = -1 Then
        ISoundStream_CloseSource
        Exit Function
    End If

    FileSeek hInput, lngSyncPos, FILE_BEGIN
    FileRead hInput, VarPtr(sync(0)), 4
    udtFrameHdr = GetFrameHeader(sync, 0)
    udtXingHdr = GetXing()

    FileSeek hInput, lngSyncPos, FILE_BEGIN

    With udtWFX
        .cbSize = MPEGLAYER3_WFX_EXTRA_BYTES
        .wFormatTag = WAVE_FORMAT_MPEGLAYER3
        .nChannels = channels
        .nAvgBytesPerSec = 128 * (1024 / 8)
        .wBitsPerSample = 0
        .nBlockAlign = 1
        .nSamplesPerSec = samplerate
        .fdwFlags = MPEGLAYER3_FLAG_PADDING_OFF
        .nBlockSize = MPEGLAYER3_BLOCK_SIZE
        .nFramesPerBlock = 1
        .nCodecDelay = 1393
        .wID = MPEGLAYER3_ID_MPEG
    End With

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

    If FileEnd(hInput) Then
        ISoundStream_Read = STREAM_END_OF_STREAM
        cBytesRead = 0
        Exit Function
    End If

    cBytesRead = FileRead(hInput, pBuffer, cBytes)
End Function

' from CDEXOS:
' http://cvs.sourceforge.net/viewcvs.py/cdexos/cdexos/ISndStreamMPEG.cpp?rev=1.17
Private Function ISoundStream_SeekTo( _
    ByVal value As Long, _
    mode As SEEK_MODE _
) As STREAM_ERR

    Dim btData()    As Byte
    Dim dblFPos     As Double
    Dim intPercent  As Integer

    If hInput.handle = INVALID_HANDLE Then
        Exit Function
    End If

    Select Case mode
        Case SEEK_PERCENT
            If value < 0 Or value > 100 Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            intPercent = value

        Case SEEK_SECONDS
            If value < 0 Or value > (ISoundInfo_Duration / 1000) Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            intPercent = value / (ISoundInfo_Duration / 1000) * 100

    End Select

    If (udtXingHdr.Flags And VBR_TOC) Then
        dblFPos = CDbl(udtXingHdr.toc(intPercent)) * CDbl(FileLength(hInput)) / 255#
    Else
        dblFPos = CDbl(intPercent) * CDbl(FileLength(hInput)) / 100#
    End If

    If CLng(dblFPos) = 0 Then dblFPos = lngSyncPos

    FileSeek hInput, CLng(dblFPos), FILE_BEGIN

    ISoundStream_SeekTo = STREAM_OK
End Function

' ***********************************
' * helpers
' ***********************************

Private Sub GetTags( _
    ByVal file As String _
)

    Dim udtID3      As modID3Tags.ID3Tag

    clsTags.ClearTags

    If Not modID3Tags.ReadID3v2(file, udtID3) Then
        If Not modID3Tags.ReadID3v1(file, udtID3) Then
            Exit Sub
        End If
    End If

    clsTags.AddTag "artist", udtID3.Artist
    clsTags.AddTag "album", udtID3.Album
    clsTags.AddTag "title", udtID3.Title
    clsTags.AddTag "genre", udtID3.Genre
    clsTags.AddTag "comment", udtID3.Comment
    clsTags.AddTag "year", udtID3.SongYear
    clsTags.AddTag "tracknumber", udtID3.TrackNr
End Sub

' look for either FF FB or FF E0
Private Function SearchFirstSync( _
    kbytes As Long, _
    unsafe As Boolean, _
    Optional startbyte As Long = 0 _
) As Long

    Dim i       As Long
    Dim sync(3) As Byte

    For i = startbyte To kbytes * 1024&
        FileSeek hInput, i, FILE_BEGIN
        FileRead hInput, VarPtr(sync(0)), 4

        If IsValidFrameHeader(GetFrameHeader(sync, 0), unsafe) Then
            SearchFirstSync = i
            Exit Function
        End If
    Next

    SearchFirstSync = -1
End Function

Private Function IsValidFrameHeader( _
    hdr As FrameHeader, _
    unsafe As Boolean _
) As Boolean

    With hdr
        If (.layer = &H0) Then Exit Function
        If (.MPEGID = &H1) Then Exit Function
        If (.sync(0) <> &HFF) Then Exit Function
        If (.Bitrate And &HF) = &HF Then Exit Function
        If unsafe Then
            If (.sync(1) And &HE0) <> &HE0 Then Exit Function
        Else
            If (.sync(1) <> &HFB) Then Exit Function
        End If
        If (.samplerate And &H3) = &H3 Then Exit Function
    End With
    IsValidFrameHeader = True
End Function

Private Function GetFrameHeader( _
    Data() As Byte, _
    offset As Byte _
) As FrameHeader

    Dim Bt  As Byte

    ' by Danielo (ActiveVB)
    With GetFrameHeader
        .sync(0) = Data(offset + 0)
        .sync(1) = Data(offset + 1)
        .sync(2) = Data(offset + 2)
        .sync(3) = Data(offset + 3)

         Bt = Data(offset + 1)
        .layer = (&H6 And Bt) / 2
        .MPEGID = (&H18 And Bt) / 8
        .ProtectBit = (&H1 And Bt)

         Bt = Data(offset + 2)
        .Bitrate = (&HF0 And Bt) / 16
        .samplerate = (&HC And Bt) / 4
        .PaddingBit = (&H2 And Bt) / 2
        .PrivateBit = (&H1 And Bt)

         Bt = Data(offset + 3)
        .emphasis = (&H3 And Bt)
        .ModeExt = (&H30 And Bt) / 16
        .CpyrightBit = (&H8 And Bt) / 8
        .OriginalBit = (&H4 And Bt) / 4
        .ChannelMode = (&HC0 And Bt) / 64
    End With
End Function

Private Function GetXing() As XingHeader

    Dim strFrame                 As String
    ' +
    ' +
    Dim lngPos                   As Long
    ' +
    ' +
    Dim btFrame()                As Byte
    ReDim btFrame(FrameSize - 1) As Byte


    FileSeek hInput, lngSyncPos, FILE_BEGIN
    FileRead hInput, VarPtr(btFrame(0)), FrameSize

    strFrame = StrConv(btFrame, vbUnicode)

    lngPos = InStr(strFrame, "Xing")
    If lngPos = 0 Then Exit Function
    lngPos = lngPos + 3

    DXCopyMemory GetXing.Flags, btFrame(lngPos), 4
    GetXing.Flags = ReverseByteOrder(GetXing.Flags)
    lngPos = lngPos + 4

    If (GetXing.Flags And VBR_FRAMES) Then
        DXCopyMemory GetXing.Frames, btFrame(lngPos), 4
        GetXing.Frames = ReverseByteOrder(GetXing.Frames)
        lngPos = lngPos + 4
    End If

    If (GetXing.Flags And VBR_BYTES) Then
        DXCopyMemory GetXing.bytes, btFrame(lngPos), 4
        GetXing.bytes = ReverseByteOrder(GetXing.bytes)
        lngPos = lngPos + 4
    End If

    If (GetXing.Flags And VBR_TOC) Then
        DXCopyMemory GetXing.toc(0), btFrame(lngPos), 100
        lngPos = lngPos + 100
    End If

    If (GetXing.Flags And VBR_SCALE) Then
        DXCopyMemory GetXing.Quality, btFrame(lngPos), 4
        GetXing.Quality = ReverseByteOrder(GetXing.Quality)
    End If

    GetXing.valid = True
End Function

Private Function ReverseByteOrder( _
    ByVal lng As Long _
) As Long

    Dim Bt(3)       As Byte
    Dim newbt(3)    As Byte

    DXCopyMemory Bt(0), lng, 4
    newbt(0) = Bt(3)
    newbt(1) = Bt(2)
    newbt(2) = Bt(1)
    newbt(3) = Bt(0)
    DXCopyMemory lng, newbt(0), 4

    ReverseByteOrder = lng
End Function

' from CDEXOS
' http://cvs.sourceforge.net/viewcvs.py/cdexos/cdexos/ISndStreamMPEG.cpp?rev=1.17
Private Property Get CurrentTime() As Long
    Dim dblRelFPos  As Double, dblPos   As Double
    Dim dA          As Double, dB       As Double
    Dim dPercent    As Double, dDelta   As Double
    ' +
    ' +
    Dim i           As Long

    If (udtXingHdr.Flags And VBR_TOC) Then

        dblRelFPos = FileSeek(hInput, 0, FILE_CURRENT) / bytes

        For i = 0 To 99
            dblPos = dblRelFPos * 255#

            If udtXingHdr.toc(i) > dblPos Then
                dDelta = udtXingHdr.toc(i) - dblPos

                dB = udtXingHdr.toc(i)

                If (i > 0) Then
                    dA = udtXingHdr.toc(i - 1)
                Else
                    dA = 0
                End If

                If (dB - dA) > 0 Then
                    dDelta = dDelta / (dB - dA)
                    dPercent = i + (1# - dDelta)
                Else
                    dPercent = i
                End If

                CurrentTime = (Frames * dPercent / 100 * SamplesPerFrame / samplerate) * 1000
                Exit Property
            End If

        Next

    Else

        Dim nFrames As Double
        nFrames = FileSeek(hInput, 0, FILE_CURRENT) / FrameSize
        CurrentTime = (CDbl(nFrames) * CDbl(SamplesPerFrame) / CDbl(samplerate)) * 1000 - 2000

    End If
End Property

Private Property Get Bitrate() As Long
    Select Case MPEGID

        Case [MPEG 1]:
            Select Case layer
                Case 1:
                    Select Case udtFrameHdr.Bitrate
                        Case &H1:   Bitrate = 32
                        Case &H2:   Bitrate = 64
                        Case &H3:   Bitrate = 96
                        Case &H4:   Bitrate = 128
                        Case &H5:   Bitrate = 160
                        Case &H6:   Bitrate = 192
                        Case &H7:   Bitrate = 224
                        Case &H8:   Bitrate = 256
                        Case &H9:   Bitrate = 288
                        Case &HA:   Bitrate = 320
                        Case &HB:   Bitrate = 352
                        Case &HC:   Bitrate = 384
                        Case &HD:   Bitrate = 416
                        Case &HE:   Bitrate = 448
                    End Select

                Case 2:
                    Select Case udtFrameHdr.Bitrate
                        Case &H1:   Bitrate = 32
                        Case &H2:   Bitrate = 48
                        Case &H3:   Bitrate = 56
                        Case &H4:   Bitrate = 64
                        Case &H5:   Bitrate = 80
                        Case &H6:   Bitrate = 96
                        Case &H7:   Bitrate = 112
                        Case &H8:   Bitrate = 128
                        Case &H9:   Bitrate = 160
                        Case &HA:   Bitrate = 192
                        Case &HB:   Bitrate = 224
                        Case &HC:   Bitrate = 256
                        Case &HD:   Bitrate = 320
                        Case &HE:   Bitrate = 384
                    End Select

                Case 3:
                    Select Case udtFrameHdr.Bitrate
                        Case &H1:   Bitrate = 32
                        Case &H2:   Bitrate = 40
                        Case &H3:   Bitrate = 48
                        Case &H4:   Bitrate = 56
                        Case &H5:   Bitrate = 64
                        Case &H6:   Bitrate = 80
                        Case &H7:   Bitrate = 96
                        Case &H8:   Bitrate = 112
                        Case &H9:   Bitrate = 128
                        Case &HA:   Bitrate = 160
                        Case &HB:   Bitrate = 192
                        Case &HC:   Bitrate = 224
                        Case &HD:   Bitrate = 256
                        Case &HE:   Bitrate = 320
                    End Select

            End Select

        Case [MPEG 2], [MPEG 2.5]:
            Select Case layer
                Case 1:
                    Select Case udtFrameHdr.Bitrate
                        Case &H1:   Bitrate = 32
                        Case &H2:   Bitrate = 48
                        Case &H3:   Bitrate = 56
                        Case &H4:   Bitrate = 64
                        Case &H5:   Bitrate = 80
                        Case &H6:   Bitrate = 96
                        Case &H7:   Bitrate = 112
                        Case &H8:   Bitrate = 128
                        Case &H9:   Bitrate = 144
                        Case &HA:   Bitrate = 160
                        Case &HB:   Bitrate = 176
                        Case &HC:   Bitrate = 192
                        Case &HD:   Bitrate = 224
                        Case &HE:   Bitrate = 256
                    End Select

                Case 2, 3:
                    Select Case udtFrameHdr.Bitrate
                        Case &H1:   Bitrate = 8
                        Case &H2:   Bitrate = 16
                        Case &H3:   Bitrate = 24
                        Case &H4:   Bitrate = 32
                        Case &H5:   Bitrate = 40
                        Case &H6:   Bitrate = 48
                        Case &H7:   Bitrate = 56
                        Case &H8:   Bitrate = 64
                        Case &H9:   Bitrate = 80
                        Case &HA:   Bitrate = 96
                        Case &HB:   Bitrate = 112
                        Case &HC:   Bitrate = 128
                        Case &HD:   Bitrate = 144
                        Case &HE:   Bitrate = 160
                    End Select

            End Select

    End Select
End Property

Private Property Get samplerate() As Long
    Select Case MPEGID

        Case [MPEG 1]:
            Select Case udtFrameHdr.samplerate
                Case 0:     samplerate = 44100
                Case 1:     samplerate = 48000
                Case 2:     samplerate = 32000
            End Select

        Case [MPEG 2]:
            Select Case udtFrameHdr.samplerate
                Case 0:     samplerate = 22050
                Case 1:     samplerate = 24000
                Case 2:     samplerate = 16000
            End Select

        Case [MPEG 2.5]:
            Select Case udtFrameHdr.samplerate
                Case 0:     samplerate = 11025
                Case 1:     samplerate = 12000
                Case 2:     samplerate = 8000
            End Select

    End Select
End Property

Private Property Get PaddingBit() As Boolean
    PaddingBit = CBool(udtFrameHdr.PaddingBit)
End Property

Private Property Get channels() As Byte
    Select Case udtFrameHdr.ChannelMode
        Case 0:     channels = 2
        Case 1:     channels = 2
        Case 2:     channels = 2
        Case 3:     channels = 1
    End Select
End Property

Private Property Get Duration() As Long
    If IsVBR Then
        Duration = (bytes / (bytes / (Frames * SamplesPerFrame / samplerate))) * 1000
    Else
        Duration = (bytes / (Bitrate * 1000) * 8) * 1000
    End If
End Property

Private Property Get Frames() As Long
    If udtXingHdr.Frames > 0 Then
        Frames = udtXingHdr.Frames
    Else
        Frames = Fix(bytes / FrameSize)
    End If
End Property

Private Property Get HeaderPosition() As Long
    HeaderPosition = lngSyncPos - 1
End Property

Private Property Get bytes() As Long
    If hInput.handle = INVALID_HANDLE Then
        Exit Property
    End If

    If udtXingHdr.bytes > 0 Then
        bytes = udtXingHdr.bytes
    Else
        bytes = FileLength(hInput) - HeaderPosition
    End If
End Property

Private Property Get IsVBR() As Boolean
    IsVBR = udtXingHdr.valid
End Property

Private Property Get SamplesPerFrame() As Integer
    Select Case MPEGID
        Case [MPEG 1]:
            Select Case layer
                Case 1:     SamplesPerFrame = 384
                Case 2, 3:  SamplesPerFrame = 1152
            End Select
        Case [MPEG 2], [MPEG 2.5]:
            Select Case layer
                Case 1:     SamplesPerFrame = 384
                Case 2:     SamplesPerFrame = 1152
                Case 3:     SamplesPerFrame = 576
            End Select
    End Select
End Property

Private Property Get FrameSize() As Long
    Dim padding As Byte

    If samplerate = 0 Then Exit Property
    If Bitrate = 0 Then Exit Property

    If PaddingBit Then
        Select Case layer
            Case 1:     padding = 4
            Case 2, 3:  padding = 1
        End Select
    End If

    If MPEGID = [MPEG 1] Then
        If layer = 1 Then
            FrameSize = Fix(((12 * (Bitrate * 1000)) / samplerate + padding) * 4)
        Else
            FrameSize = Fix((144 * (Bitrate * 1000)) / samplerate + padding)
        End If
    Else
        If layer = 1 Then
            FrameSize = Fix(((24 * (Bitrate * 1000)) / samplerate + padding) * 4)
        Else
            FrameSize = Fix((72 * (Bitrate * 1000)) / samplerate + padding)
        End If
    End If
End Property

Private Property Get MPEGID() As eMPEGID
    MPEGID = udtFrameHdr.MPEGID
End Property

Private Property Get layer() As Byte
    Select Case udtFrameHdr.layer
        Case 1:     layer = 3
        Case 2:     layer = 2
        Case 3:     layer = 1
    End Select
End Property

Private Property Get VBRBitrate() As Long
    If Not udtXingHdr.valid Then Exit Property
    VBRBitrate = CLng(udtXingHdr.bytes / udtXingHdr.Frames * samplerate / (IIf(MPEGID = [MPEG 1], 144000, 72000)))
End Property
