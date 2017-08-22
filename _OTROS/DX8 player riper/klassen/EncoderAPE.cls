VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderAPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Monkey's Audio Encoder
'
' lossless PCM compression

Implements IEncoder

Private Declare Function APECompress_Create Lib "MACDll.dll" _
Alias "c_APECompress_Create" ( _
    pErrorCode As Long _
) As Long

Private Declare Sub APECompress_Destroy Lib "MACDll.dll" _
Alias "c_APECompress_Destroy" ( _
    ByVal handle As APE_HANDLE _
)

Private Declare Function APECompress_Start Lib "MACDll.dll" _
Alias "c_APECompress_Start" ( _
    ByVal handle As APE_HANDLE, _
    ByVal pOutputFilename As String, _
    pwfeInput As MMWAVEFORMATEX, _
    ByVal nMaxAudioBytes As Long, _
    ByVal nCompressionLevel As Long, _
    ByVal pHeaderData As Long, _
    ByVal nHeaderBytes As Long _
) As Long

Private Declare Function APECompress_AddData Lib "MACDll.dll" _
Alias "c_APECompress_AddData" ( _
    ByVal handle As APE_HANDLE, _
    pData As Any, _
    ByVal nBytes As Long _
) As Long

Private Declare Function APECompress_GetBufferBytesAvailable Lib "MACDll.dll" _
Alias "c_APECompress_GetBufferBytesAvailable" ( _
    ByVal handle As APE_HANDLE _
) As Long

Private Declare Function APECompress_LockBuffer Lib "MACDll.dll" _
Alias "c_APECompress_LockBuffer" ( _
    ByVal handle As APE_HANDLE, _
    pBytesAvailable As Long _
) As Long

Private Declare Function APECompress_UnlockBuffer Lib "MACDll.dll" _
Alias "c_APECompress_UnlockBuffer" ( _
    ByVal handle As APE_HANDLE, _
    ByVal nBytesAdded As Long, _
    ByVal bProcess As Long _
) As Long

Private Declare Function APECompress_Finish Lib "MACDll.dll" _
Alias "c_APECompress_Finish" ( _
    ByVal handle As APE_HANDLE, _
    ByVal pTerminatingData As Long, _
    ByVal nTerminatingBytes As Long, _
    ByVal nWAVTerminatingBytes As Long _
) As Long

Private Declare Function APECompress_Kill Lib "MACDll.dll" _
Alias "c_APECompress_Kill" ( _
    ByVal handle As APE_HANDLE _
) As Long

Private Type APE_TAG_FOOTER
    preample        As String * 8
    tagver          As Long
    tagsize         As Long
    cTags           As Long
    Flags           As Long
    zero1           As Long
    zero2           As Long
End Type

Private Type APE_TAG_ITEM
    size            As Long
    Flags           As Long
End Type

Public Enum APE_COMPRESSION_LEVELS
    COMPRESSION_LEVEL_FAST = 1000
    COMPRESSION_LEVEL_NORMAL = 2000
    COMPRESSION_LEVEL_HIGH = 3000
    COMPRESSION_LEVEL_EXTRA_HIGH = 4000
    COMPRESSION_LEVEL_INSANE = 5000
End Enum

Private Enum APE_HANDLE
    INVALID_HANDLE = 0
End Enum

Private Const MAX_AUDIO_BYTES_UNKNOWN               As Long = -1&
Private Const CREATE_WAV_HEADER_ON_DECOMPRESSION    As Long = -1&

Private Const BytesPerEncode As Long = 8& * 1024&

Private hAPE            As APE_HANDLE
Private udeCompLevel    As APE_COMPRESSION_LEVELS

Private strTitle        As String
Private strAlbum        As String
Private strArtist       As String

Private strDest         As String

' ********************************
' * main functions
' ********************************

Private Function IEncoder_DeInit() As STREAM_ERR
    Dim lngRet  As Long

    lngRet = APECompress_Finish(hAPE, 0, 0, 0)
    APECompress_Destroy hAPE
    hAPE = 0

    WriteAPETags strDest, strArtist, _
                          strAlbum, _
                          strTitle

    IEncoder_DeInit = STREAM_OK
End Function

Private Function IEncoder_Encode( _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As STREAM_ERR

    Dim lngRet  As Long

    If bytes > BytesPerEncode Then
        IEncoder_Encode = STREAM_OUT_OF_RANGE
        Exit Function
    End If

    lngRet = APECompress_AddData(hAPE, _
                                 ByVal ptr, _
                                 bytes)

    If lngRet = 0 Then
        IEncoder_Encode = STREAM_OK
    Else
        IEncoder_Encode = STREAM_NOT_SUPPORTED
    End If
End Function

Private Function IEncoder_Init( _
    ByVal dstFilename As String, _
    ByVal srcSampleRate As Long, _
    ByVal srcChannels As Integer, _
    ByVal srcBitsPerSample As Integer, _
    cbPCMData As Long _
) As STREAM_ERR

    Dim lngRet      As Long
    Dim wfx         As MMWAVEFORMATEX

    IEncoder_Init = STREAM_NOT_SUPPORTED
    If dstFilename = "" Then Exit Function
    If srcSampleRate <= 0 Then Exit Function
    If srcChannels <= 0 Then Exit Function
    If srcBitsPerSample <= 0 Then Exit Function

    strDest = dstFilename

    If hAPE <> 0 Then IEncoder_DeInit

    wfx = CreateWFX(srcSampleRate, _
                    srcChannels, _
                    srcBitsPerSample)

    hAPE = APECompress_Create(lngRet)
    If lngRet <> 0 Then
        IEncoder_Init = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    lngRet = APECompress_Start(hAPE, _
                               strDest, _
                               wfx, _
                               MAX_AUDIO_BYTES_UNKNOWN, _
                               udeCompLevel, _
                               0, _
                               CREATE_WAV_HEADER_ON_DECOMPRESSION)

    If lngRet = 0 Then
        cbPCMData = BytesPerEncode
        IEncoder_Init = STREAM_OK
    Else
        IEncoder_Init = STREAM_NOT_SUPPORTED
    End If
End Function

' *******************************
' * Encoder settings
' *******************************

Private Property Get IEncoder_Extension() As String
    IEncoder_Extension = "ape"
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

Public Property Get CompressionLevel() As APE_COMPRESSION_LEVELS
    CompressionLevel = udeCompLevel
End Property

Public Property Let CompressionLevel( _
    ByVal newval As APE_COMPRESSION_LEVELS _
)

    udeCompLevel = newval
End Property

' *******************************
' Helpers
' *******************************

Private Sub WriteAPETags( _
    ByVal FileName As String, _
    ByVal Artist As String, _
    ByVal Album As String, _
    ByVal Title As String _
)

    Dim lngTagsSize As Long
    Dim lngTagStart As Long
    Dim cTags       As Long
    Dim udtFooter   As APE_TAG_FOOTER
    Dim udtItem     As APE_TAG_ITEM

    Dim fh          As Integer

    fh = FreeFile
    Open FileName For Binary As #fh
        Seek #fh, LOF(fh)
        lngTagStart = Seek(fh)

        udtItem.size = Len(strArtist) + 1
        Put #fh, , udtItem
        Put #fh, , CStr("Artist" & Chr$(0))
        Put #fh, , CStr(Artist & Chr$(0))
        cTags = cTags + 1

        udtItem.size = Len(strAlbum) + 1
        Put #fh, , udtItem
        Put #fh, , CStr("Album" & Chr$(0))
        Put #fh, , CStr(Album & Chr$(0))
        cTags = cTags + 1

        udtItem.size = Len(strTitle) + 1
        Put #fh, , udtItem
        Put #fh, , CStr("Title" & Chr$(0))
        Put #fh, , CStr(Title & Chr$(0))
        cTags = cTags + 1

        lngTagsSize = Seek(fh) - lngTagStart + Len(udtFooter)

        With udtFooter
            .preample = "APETAGEX"
            .cTags = cTags
            .tagsize = lngTagsSize
            .tagver = 1000
        End With

        Put #fh, , udtFooter
    Close #fh
End Sub
