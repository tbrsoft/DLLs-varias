VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderWAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' WAV Encoder with ACM support

Implements IEncoder

' Vorbis ACM has really big WFXs.
' so just in case...
Private Const WFXSize   As Long = 128& * 1024&

Private Const WAV_RIFF  As Long = 1179011410
Private Const WAV_WAVE  As Long = 1163280727
Private Const WAV_DATA  As Long = 1635017060
Private Const WAV_FMT   As Long = 544501094

Private Type WAVRIFF
    RIFF                As Long         ' "RIFF"
    hdrlen              As Long
    WAVE                As Long         ' "WAVE"
    fmt                 As Long         ' "fmt "
    chnksize            As Long
End Type

Private btWfxOut()      As Byte
Private lngWFXOutLen    As Long
Private strFormatID     As String
Private strFormatTag    As String

Private btInp()         As Byte
Private btOut()         As Byte
Private lngInpLen       As Long
Private lngOutLen       As Long

Private udtStreamHdr    As ACMSTREAMHEADER
Private hStream         As HACMSTREAM

Private strFile         As String
Private hFOut           As hFile

Private strArtist       As String
Private strAlbum        As String
Private strTitle        As String

Private blnWriteHdr     As Boolean

' *************************************
' * constructor
' *************************************

Private Sub Class_Initialize()
    Dim wfx As MMWAVEFORMATEX

    ReDim btWfxOut(WFXSize - 1) As Byte
    lngInpLen = 16& * 1024&
    blnWriteHdr = True

    wfx = CreateWFX(44100, 2, 16)
    SetFormat VarPtr(wfx), Len(wfx), "44.1 kHz 16 Bit Stereo", "PCM"
End Sub

' *************************************
' * Encoder settings
' *************************************

Private Property Get IEncoder_Title() As String
    IEncoder_Title = strTitle
End Property

Private Property Let IEncoder_Title( _
    ByVal RHS As String _
)

    strTitle = RHS
End Property

Private Property Let IEncoder_Album( _
    ByVal RHS As String _
)

    strAlbum = RHS
End Property

Private Property Get IEncoder_Album() As String
    IEncoder_Album = strAlbum
End Property

Private Property Let IEncoder_Artist( _
    ByVal RHS As String _
)

    strArtist = RHS
End Property

Private Property Get IEncoder_Artist() As String
    IEncoder_Artist = strArtist
End Property

Private Property Get IEncoder_Extension() As String
    IEncoder_Extension = "wav"
End Property

Public Property Get WriteWAVHeader() As Boolean
    WriteWAVHeader = blnWriteHdr
End Property

Public Property Let WriteWAVHeader( _
    ByVal bln As Boolean _
)

    blnWriteHdr = bln
End Property

Public Property Get FormatTag() As String
    FormatTag = strFormatTag
End Property

Public Property Get FormatID() As String
    FormatID = strFormatID
End Property

Public Sub SetFormat( _
    ByVal ptr As Long, _
    ByVal bytes As Long, _
    Optional strFmtID As String, _
    Optional strFmtTag As String _
)

    If 0 = IsBadReadPtr(ByVal ptr, bytes) Then
        If bytes > UBound(btWfxOut) + 1 Then
            ReDim btWfxOut(bytes - 1) As Byte
        End If
        DXCopyMemory btWfxOut(0), ByVal ptr, bytes

        lngWFXOutLen = bytes

        strFormatTag = strFmtTag
        strFormatID = strFmtID
    End If
End Sub

Public Function SelectFormat( _
    ByVal samplerate As Long, _
    ByVal channels As Integer, _
    ByVal bitspersample As Integer, _
    Optional ByVal hwndOwner As Long, _
    Optional ByVal Title As String = "WAV Format" _
) As STREAM_ERR

    Dim btTitle()   As Byte
    Dim btWFX()     As Byte
    Dim wfx         As MMWAVEFORMATEX
    Dim wfxinp      As MMWAVEFORMATEX
    Dim udtFmtCh    As ACMFORMATCHOOSEA

    ReDim btWFX(WFXSize - 1) As Byte

    btTitle = StrConv(Title & Chr$(0), vbFromUnicode)

    With wfxinp
        .nSamplesPerSec = samplerate
        .nChannels = channels
        .wBitsPerSample = bitspersample
        .wFormatTag = WAVE_FORMAT_PCM
        .nBlockAlign = .nChannels * (.wBitsPerSample / 8)
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
    End With

    DXCopyMemory btWFX(0), wfxinp, Len(wfxinp)

    With udtFmtCh
        .cbStruct = LenB(udtFmtCh)
        .hwndOwner = hwndOwner
        .pwfx = VarPtr(btWFX(0))
        .cbwfx = WFXSize
        .pwfxEnum = VarPtr(wfxinp)
        .fdwEnum = ACM_FORMATENUMF_CONVERT
        .pszTitle = VarPtr(btTitle(0))
    End With

    If acmFormatChoose(udtFmtCh) = 0 Then
        btWfxOut = btWFX
        DXCopyMemory wfx, btWFX(0), Len(wfx)
        lngWFXOutLen = CLng(18) + CLng(wfx.cbSize)
        strFormatID = RemNullChars(udtFmtCh.szFormat)
        strFormatTag = RemNullChars(udtFmtCh.szFormatTag)
        SelectFormat = STREAM_OK
    Else
        SelectFormat = STREAM_CANCELED
    End If
End Function

' *************************************
' * main functions
' *************************************

Private Function IEncoder_DeInit() As STREAM_ERR
    ' ACM Flush
    udtStreamHdr.cbSrcLength = lngInpLen
    acmStreamConvert hStream, udtStreamHdr, ACM_STREAMCONVERTF_END
    FileWrite hFOut, udtStreamHdr.pbDst, udtStreamHdr.cbDstLengthUsed
    If blnWriteHdr Then WriteWAVHdr
    FileClose hFOut

    acmStreamUnprepareHeader hStream, udtStreamHdr, 0
    acmStreamClose hStream, 0: hStream = 0

    IEncoder_DeInit = STREAM_OK
End Function

Private Function IEncoder_Encode( _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As STREAM_ERR

    Dim lngRet As Long

    If bytes > lngInpLen Then
        IEncoder_Encode = STREAM_TOO_BIG
        Exit Function
    End If

    If 0 = IsBadReadPtr(ByVal ptr, bytes) Then
        DXCopyMemory btInp(0), ByVal ptr, bytes
    End If

    udtStreamHdr.cbSrcLength = bytes

    lngRet = acmStreamConvert(hStream, _
                              udtStreamHdr, _
                              ACM_STREAMCONVERTF_BLOCKALIGN)

    If lngRet <> 0 Then
        If udtStreamHdr.cbSrcLengthUsed = 0 Then
            IEncoder_Encode = STREAM_NEED_MORE
        Else
            IEncoder_Encode = STREAM_OK
        End If
    Else
        FileWrite hFOut, udtStreamHdr.pbDst, udtStreamHdr.cbDstLengthUsed
        IEncoder_Encode = STREAM_OK
    End If
End Function

Private Function IEncoder_Init( _
    ByVal dstFilename As String, _
    ByVal srcSampleRate As Long, _
    ByVal srcChannels As Integer, _
    ByVal srcBitsPerSample As Integer, _
    cbPCMData As Long _
) As STREAM_ERR

    Dim wfxinp      As MMWAVEFORMATEX
    Dim btHdr()     As Byte
    Dim lngRet      As Long

    wfxinp = CreateWFX(srcSampleRate, _
                       srcChannels, _
                       srcBitsPerSample)

    lngRet = acmStreamOpen(hStream, _
                           0, _
                           wfxinp, _
                           btWfxOut(0), _
                           0, 0, 0, 0)

    If lngRet <> 0 Then
        IEncoder_Init = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    lngRet = acmStreamSize(hStream, _
                           lngInpLen, _
                           lngOutLen, _
                           ACM_STREAMSIZEF_SOURCE)

    If lngRet <> 0 Then
        IEncoder_Init = STREAM_NOT_SUPPORTED
        acmStreamClose hStream, 0
        Exit Function
    End If

    ReDim btInp(lngInpLen - 1) As Byte
    ReDim btOut(lngOutLen - 1) As Byte

    With udtStreamHdr
        .cbStruct = Len(udtStreamHdr)
        .cbSrcLength = Len(udtStreamHdr)
        .cbDstLength = lngOutLen
        .cbSrcLength = lngInpLen
        .pbDst = VarPtr(btOut(0))
        .pbSrc = VarPtr(btInp(0))
    End With

    lngRet = acmStreamPrepareHeader(hStream, _
                                    udtStreamHdr, _
                                    0)

    If lngRet <> 0 Then
        IEncoder_Init = STREAM_NOT_SUPPORTED
        acmStreamClose hStream, 0
        Exit Function
    End If

    hFOut = FileOpen(dstFilename, _
                     GENERIC_WRITE, _
                     FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                     CREATE_ALWAYS)

    If hFOut.handle = INVALID_HANDLE Then
        IEncoder_Init = STREAM_INVALID_SOURCE
        acmStreamUnprepareHeader hStream, udtStreamHdr, 0
        acmStreamClose hStream, 0
        Exit Function
    End If

    If blnWriteHdr Then
        ' space for the WAV headers
        ReDim btHdr(20 + lngWFXOutLen + 4 * 2 - 1) As Byte
        FileWrite hFOut, VarPtr(btHdr(0)), UBound(btHdr) + 1
    End If

    cbPCMData = lngInpLen

    IEncoder_Init = STREAM_OK
End Function

' *************************************
' * helpers
' *************************************

Private Sub WriteWAVHdr()
    Dim udtHdr      As WAVRIFF
    Dim lngDataID   As Long
    Dim lngData     As Long

    With udtHdr
        .RIFF = WAV_RIFF
        .WAVE = WAV_WAVE
        .fmt = WAV_FMT
        .hdrlen = FileSeek(hFOut, 0, FILE_END) - 8
        .chnksize = lngWFXOutLen
    End With

    FileSeek hFOut, 0, FILE_BEGIN
    FileWrite hFOut, VarPtr(udtHdr), Len(udtHdr)

    FileWrite hFOut, VarPtr(btWfxOut(0)), lngWFXOutLen

    lngDataID = WAV_DATA
    FileWrite hFOut, VarPtr(lngDataID), 4

    lngData = (udtHdr.hdrlen + 8) - (Len(udtHdr) - lngWFXOutLen) - 4 * 2
    FileWrite hFOut, VarPtr(lngData), 4
End Sub
