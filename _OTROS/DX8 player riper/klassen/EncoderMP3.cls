VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' MPEG Layer 3 Encoder with Lame (lame_enc.dll)

Implements IEncoder

Private Enum HBE_STREAM
    INVALID_HANDLE = 0
End Enum

Private Enum BE_CONFIG
    BE_CONFIG_MP3 = 0
    BE_CONFIG_LAME = 256
End Enum

Private Enum BE_ERR
    BE_ERR_SUCCESSFUL = &H0
    BE_ERR_INVALID_FORMAT = &H1
    BE_ERR_INVALID_FORMAT_PARAMETERS = &H2
    BE_ERR_NO_MORE_HANDLES = &H3
    BE_ERR_INVALID_HANDLE = &H4
    BE_ERR_BUFFER_TOO_SMALL = &H5
End Enum

Private Enum BE_CHANNELS
    BE_MP3_MODE_STEREO = 0
    BE_MP3_MODE_JSTEREO = 1
    BE_MP3_MODE_DUALCHANNEL = 2
    BE_MP3_MODE_MONO = 3
End Enum

Private Enum MPEG_VER
    MPEG1 = 1
    MPEG2 = 0
End Enum

Private Const CURRENT_STRUCT_VERSION    As Long = 1
Private Const CURRENT_STRUCT_SIZE       As Long = 331

Private Const MP3_QUALITY               As Long = 8

Public Enum VBR_QUALITY
    VBR_QUALITY_0_HIGHEST = 0
    VBR_QUALITY_1_HIGH = 1
    VBR_QUALITY_2_HIGH = 2
    VBR_QUALITY_3_MID = 3
    VBR_QUALITY_4_MID = 4
    VBR_QUALITY_5_MID = 5
    VBR_QUALITY_6_MID = 6
    VBR_QUALITY_7_LOW = 7
    VBR_QUALITY_8_LOW = 8
    VBR_QUALITY_9_LOWEST = 9
End Enum

Private Enum VBRMethod
    VBR_METHOD_NONE = -1
    VBR_METHOD_DEFAULT = 0
    VBR_METHOD_OLD = 1
    VBR_METHOD_NEW = 2
    VBR_METHOD_MTRH = 3
    VBR_METHOD_ABR = 4
End Enum

Private Enum LAME_QUALITY_PRESET
    LQP_NOPRESET = -1

    LQP_NORMAL_QUALITY = 0
    LQP_LOW_QUALITY = 1
    LQP_HIGH_QUALITY = 2
    LQP_VOICE_QUALITY = 3
    LQP_R3MIX = 4
    LQP_VERYHIGH_QUALITY = 5
    LQP_STANDARD = 6
    LQP_FAST_STANDARD = 7
    LQP_EXTREME = 8
    LQP_FAST_EXTREME = 9
    LQP_INSANE = 10
    LQP_ABR = 11
    LQP_CBR = 12
    LQP_MEDIUM = 13
    LQP_FAST_MEDIUM = 14

    LQP_PHONE = 1000
    LQP_SW = 2000
    LQP_AM = 3000
    LQP_FM = 4000
    LQP_VOICE = 5000
    LQP_RADIO = 6000
    LQP_TAPE = 7000
    LQP_HIFI = 8000
    LQP_CD = 9000
    LQP_STUDIO = 10000
End Enum

Private Type BE_CONFIG_LHV1
    dwConfig            As Long

    dwStructVersion     As Long
    dwStructSize        As Long
    dwSampleRate        As Long
    dwReSampleRate      As Long
    nMode               As Long
    dwBitRate           As Long
    dwMaxBitrate        As Long
    nPreset             As Long
    dwMpegVersion       As Long
    dwPsyModel          As Long
    dwEmphasis          As Long

    bPrivate            As Long
    bCRC                As Long
    bCopyright          As Long
    bOriginal           As Long

    bWriteVBRHeader     As Long
    bEnableVBR          As Long
    nVBRQuality         As Long
    dwVbrAbr_bps        As Long
    nVBRMethod          As Long
    bNoRes              As Long

    bStrictIso          As Long
    nQuality            As Integer

    bPadding(236)       As Byte
End Type

Private Type BE_CONFIG_AAC
    dwConfig            As Long

    dwSampleRate        As Long
    byMode              As Byte
    wBitrate            As Integer
    byEncodingMethod    As Byte
End Type

Private Type BE_CONFIG_MP3
    dwConfig            As Long

    dwSampleRate        As Long
    byMode              As Byte
    wBitrate            As Integer
    bPrivate            As Long
    bCRC                As Long
    bCopyright          As Long
    bOriginal           As Long
End Type

Private Const BE_MAX_HOMEPAGE       As Long = 128

Private Type BE_VERSION
    byDLLMajorVersion   As Byte
    byDLLMinorVersion   As Byte
    byMajorVersion      As Byte
    byMinorVersion      As Byte
    byDay               As Byte
    byMonth             As Byte
    wYear               As Integer
    zHomepage           As String * BE_MAX_HOMEPAGE
    byAlphaLevel        As Byte
    byBetaLevel         As Byte
    byMMXEnabled        As Byte
    btReserved(126)     As Byte
End Type

Private clsLame         As clsDLL

Private strArtist       As String
Private strAlbum        As String
Private strTitle        As String

Private udtConfig       As BE_CONFIG_LHV1

Private hStream         As HBE_STREAM
Private hOut            As hFile

Private dwMP3Buffer     As Long
Private btMP3Data()     As Byte

Private strFile         As String

Private blnLameReady    As Boolean

' ***************************************
' * constructor
' ***************************************

Private Sub Class_Initialize()
    Set clsLame = New clsDLL

    blnLameReady = clsLame.LoadDLL("lame_enc.dll")
    clsLame.IsCDECL = True
End Sub

Private Sub Class_Terminate()
    clsLame.UnloadDLL
    Set clsLame = Nothing
End Sub

' ***************************************
' * Encoder settings
' ***************************************

Public Property Let VBRQuality( _
    ByVal q As VBR_QUALITY _
)

    udtConfig.nVBRQuality = q
End Property

Public Property Get VBRQuality() As VBR_QUALITY
    VBRQuality = udtConfig.nVBRQuality
End Property

Public Property Let VBRMaxBitrate( _
    ByVal bit As Long _
)

    udtConfig.dwMaxBitrate = bit
End Property

Public Property Get VBRMaxBitrate() As Long
    VBRMaxBitrate = udtConfig.dwMaxBitrate
End Property

Public Property Let VBR( _
    ByVal bln As Boolean _
)

    udtConfig.bEnableVBR = Abs(bln)
End Property

Public Property Get VBR() As Boolean
    VBR = CBool(udtConfig.bEnableVBR)
End Property

Public Property Let Bitrate( _
ByVal bit As Long _
)

    udtConfig.dwBitRate = bit
End Property

Public Property Get Bitrate() As Long
    Bitrate = udtConfig.dwBitRate
End Property

Private Property Let IEncoder_Title( _
    ByVal RHS As String _
)

    strTitle = RHS
End Property

Private Property Get IEncoder_Title() As String
    IEncoder_Title = strTitle
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
    IEncoder_Extension = "mp3"
End Property

' ***************************************
' * main functions
' ***************************************

Private Function IEncoder_DeInit() As STREAM_ERR
    Dim dwData  As Long

    beDeinitStream hStream, VarPtr(btMP3Data(0)), dwData
    FileWrite hOut, VarPtr(btMP3Data(0)), dwData

    beCloseStream hStream
    hStream = 0

    FileClose hOut
    hOut.handle = 0

    beWriteVBRHeader strFile
    WriteID3Tags strFile

    IEncoder_DeInit = STREAM_OK
End Function

Private Function IEncoder_Encode( _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As STREAM_ERR

    Dim ret As BE_ERR

    ret = beEncodeChunk(hStream, _
                        bytes / 2, _
                        ptr, _
                        VarPtr(btMP3Data(0)), _
                        bytes)

    If ret <> BE_ERR_SUCCESSFUL Then
        IEncoder_Encode = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    FileWrite hOut, VarPtr(btMP3Data(0)), bytes
    IEncoder_Encode = STREAM_OK
End Function

Private Function IEncoder_Init( _
    ByVal dstFilename As String, _
    ByVal srcSampleRate As Long, _
    ByVal srcChannels As Integer, _
    ByVal srcBitsPerSample As Integer, _
    cbPCMData As Long _
) As STREAM_ERR

    Dim ret As BE_ERR

    IEncoder_Init = STREAM_NOT_SUPPORTED

    If Not blnLameReady Then Exit Function

    If dstFilename = "" Then Exit Function
    If srcBitsPerSample <> 16 Then Exit Function
    If srcChannels > 2 Then Exit Function
    If srcSampleRate < 1 Then Exit Function

    If hStream <> 0 Then IEncoder_DeInit

    With udtConfig
        .dwConfig = BE_CONFIG_LAME
        .dwStructVersion = CURRENT_STRUCT_VERSION
        .dwStructSize = CURRENT_STRUCT_SIZE
        .dwSampleRate = srcSampleRate
        If srcChannels = 1 Then
            .nMode = BE_MP3_MODE_MONO
        Else
            .nMode = BE_MP3_MODE_JSTEREO
        End If
        .bWriteVBRHeader = 1
        .nVBRMethod = VBR_METHOD_NEW
        .bNoRes = 1
        .nQuality = (((Not MP3_QUALITY) And &HFF) * &H100&) Or MP3_QUALITY
    End With

    hOut = FileOpen(dstFilename, GENERIC_WRITE, , CREATE_ALWAYS)
    If hOut.handle = INVALID_HANDLE Then
        IEncoder_Init = STREAM_INVALID_SOURCE
        Exit Function
    End If

    ret = beInitStream(VarPtr(udtConfig), _
                       cbPCMData, _
                       dwMP3Buffer, _
                       hStream)

    If ret <> BE_ERR_SUCCESSFUL Then
        IEncoder_Init = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    ReDim btMP3Data(dwMP3Buffer - 1) As Byte
    strFile = dstFilename

    IEncoder_Init = STREAM_OK
End Function

' ***************************************
' * Lame API Wrapper
' ***************************************

Private Function beWriteVBRHeader( _
    ByVal strFile As String _
) As Long

    Dim btFile() As Byte
    btFile = StrConv(strFile & Chr$(0), vbFromUnicode)

    beWriteVBRHeader = clsLame.CallFunc("beWriteVBRHeader", _
                                        VarPtr(btFile(0)))
End Function

Private Function beCloseStream( _
    ByVal hbeStream As HBE_STREAM _
) As Long

    beCloseStream = clsLame.CallFunc("beCloseStream", _
                                     hbeStream)
End Function

Private Function beDeinitStream( _
    ByVal hbeStream As HBE_STREAM, _
    ByVal pOutput As Long, _
    pdwOutput As Long _
) As Long

    beDeinitStream = clsLame.CallFunc("beDeinitStream", _
                                      hbeStream, _
                                      pOutput, _
                                      VarPtr(pdwOutput))
End Function

Private Function beEncodeChunk( _
    ByVal hbeStream As HBE_STREAM, _
    ByVal nSamples As Long, _
    ByVal pSamples As Long, _
    ByVal pOutput As Long, _
    pdwOutput As Long _
) As Long

    beEncodeChunk = clsLame.CallFunc("beEncodeChunk", _
                                     hbeStream, _
                                     nSamples, _
                                     pSamples, _
                                     pOutput, _
                                     VarPtr(pdwOutput))
End Function

Private Function beVersion() As BE_VERSION
    Dim ver As BE_VERSION

    clsLame.CallFunc "beVersion", VarPtr(ver)
    beVersion = ver
End Function

Private Function beInitStream( _
    ByVal confptr As Long, _
    pdwSamples As Long, _
    pdwBufferSize As Long, _
    phbeStream As HBE_STREAM _
) As Long

    beInitStream = clsLame.CallFunc("beInitStream", _
                                    confptr, _
                                    VarPtr(pdwSamples), _
                                    VarPtr(pdwBufferSize), _
                                    VarPtr(phbeStream))
End Function

' ***************************************
' * helpers
' ***************************************

Private Sub WriteID3Tags( _
    ByVal mp3file As String _
)

    Dim udtID3  As ID3Tag

    With udtID3
        .Artist = strArtist
        .Album = strAlbum
        .Title = strTitle
    End With

    modID3Tags.WriteID3v1 mp3file, udtID3, True
    modID3Tags.WriteID3v2 mp3file, udtID3, VERSION_2_3, True, True
End Sub
