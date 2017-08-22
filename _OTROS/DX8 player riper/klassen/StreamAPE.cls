VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamAPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Monkey's Audio Decoder (MACDll.dll)

Implements ISoundStream
Implements ISoundInfo

Private Declare Function APEDecompress_Create Lib "MACDll.dll" _
Alias "c_APEDecompress_Create" ( _
    ByVal pFilename As String, _
    ByVal pErrorCode As Long _
) As APE_HANDLE

Private Declare Sub APEDecompress_Destroy Lib "MACDll.dll" _
Alias "c_APEDecompress_Destroy" ( _
    ByVal handle As APE_HANDLE _
)

Private Declare Function APEDecompress_GetInfo Lib "MACDll.dll" _
Alias "c_APEDecompress_GetInfo" ( _
    ByVal handle As APE_HANDLE, _
    ByVal Info As APE_DECOMPRESS_FIELDS, _
    ByVal nParam1 As Long, _
    ByVal nParam2 As Long _
) As Long

Private Declare Function APEDecompress_Seek Lib "MACDll.dll" _
Alias "c_APEDecompress_Seek" ( _
    ByVal handle As APE_HANDLE, _
    ByVal nBlockOffset As Long _
) As Long

Private Declare Function APEDecompress_GetData Lib "MACDll.dll" _
Alias "c_APEDecompress_GetData" ( _
    ByVal handle As APE_HANDLE, _
    ByVal pBuffer As Long, _
    ByVal intBlocks As Long, _
    ByRef pBlocksRetrieved As Long _
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

Private Enum APE_HANDLE
    ape_invalid_handle = 0
End Enum

Private Enum APE_COMPRESSION_LEVELS
    COMPRESSION_LEVEL_FAST = 1000
    COMPRESSION_LEVEL_NORMAL = 2000
    COMPRESSION_LEVEL_HIGH = 3000
    COMPRESSION_LEVEL_EXTRA_HIGH = 4000
    COMPRESSION_LEVEL_INSANE = 5000
End Enum

Private Enum APE_DECOMPRESS_FIELDS
    APE_INFO_FILE_VERSION = 1000            ' version of the APE file * 1000 (3.93 = 3930) [ignored, ignored]
    APE_INFO_COMPRESSION_LEVEL = 1001       ' compression level of the APE file [ignored, ignored]
    APE_INFO_FORMAT_FLAGS = 1002            ' format flags of the APE file [ignored, ignored]
    APE_INFO_SAMPLE_RATE = 1003             ' sample rate (Hz) [ignored, ignored]
    APE_INFO_BITS_PER_SAMPLE = 1004         ' bits per sample [ignored, ignored]
    APE_INFO_BYTES_PER_SAMPLE = 1005        ' number of bytes per sample [ignored, ignored]
    APE_INFO_CHANNELS = 1006                ' channels [ignored, ignored]
    APE_INFO_BLOCK_ALIGN = 1007             ' block alignment [ignored, ignored]
    APE_INFO_BLOCKS_PER_FRAME = 1008        ' number of blocks in a frame (frames are used internally)  [ignored, ignored]
    APE_INFO_FINAL_FRAME_BLOCKS = 1009      ' blocks in the final frame (frames are used internally) [ignored, ignored]
    APE_INFO_TOTAL_FRAMES = 1010            ' total number frames (frames are used internally) [ignored, ignored]
    APE_INFO_WAV_HEADER_BYTES = 1011        ' header bytes of the decompressed WAV [ignored, ignored]
    APE_INFO_WAV_TERMINATING_BYTES = 1012   ' terminating bytes of the decompressed WAV [ignored, ignored]
    APE_INFO_WAV_DATA_BYTES = 1013          ' data bytes of the decompressed WAV [ignored, ignored]
    APE_INFO_WAV_TOTAL_BYTES = 1014         ' total bytes of the decompressed WAV [ignored, ignored]
    APE_INFO_APE_TOTAL_BYTES = 1015         ' total bytes of the APE file [ignored, ignored]
    APE_INFO_TOTAL_BLOCKS = 1016            ' total blocks of audio data [ignored, ignored]
    APE_INFO_LENGTH_MS = 1017               ' length in ms (1 sec = 1000 ms) [ignored, ignored]
    APE_INFO_AVERAGE_BITRATE = 1018         ' average bitrate of the APE [ignored, ignored]
    APE_INFO_FRAME_BITRATE = 1019           ' bitrate of specified APE frame [frame index, ignored]
    APE_INFO_DECOMPRESSED_BITRATE = 1020    ' bitrate of the decompressed WAV [ignored, ignored]
    APE_INFO_PEAK_LEVEL = 1021              ' peak audio level (obsolete) (-1 is unknown) [ignored, ignored]
    APE_INFO_SEEK_BIT = 1022                ' bit offset [frame index, ignored]
    APE_INFO_SEEK_BYTE = 1023               ' byte offset [frame index, ignored]
    APE_INFO_WAV_HEADER_DATA = 1024         ' error code [buffer *, max bytes]
    APE_INFO_WAV_TERMINATING_DATA = 1025    ' error code [buffer *, max bytes]
    APE_INFO_WAVEFORMATEX = 1026            ' error code [waveformatex *, ignored]
    APE_INFO_IO_SOURCE = 1027               ' I/O source (CIO *) [ignored, ignored]
    APE_INFO_FRAME_BYTES = 1028             ' bytes (compressed) of the frame [frame index, ignored]
    APE_INFO_FRAME_BLOCKS = 1029            ' blocks in a given frame [frame index, ignored]
    APE_INFO_TAG = 1030                     ' point to tag (CAPETag *) [ignored, ignored]

    APE_DECOMPRESS_CURRENT_BLOCK = 2000     ' current block location [ignored, ignored]
    APE_DECOMPRESS_CURRENT_MS = 2001        ' current millisecond location [ignored, ignored]
    APE_DECOMPRESS_TOTAL_BLOCKS = 2002      ' total blocks in the decompressors range [ignored, ignored]
    APE_DECOMPRESS_LENGTH_MS = 2003         ' total blocks in the decompressors range [ignored, ignored]
    APE_DECOMPRESS_CURRENT_BITRATE = 2004   ' current bitrate [ignored, ignored]
    APE_DECOMPRESS_AVERAGE_BITRATE = 2005   ' average bitrate (works with ranges) [ignored, ignored]

    APE_INTERNAL_INFO = 3000                ' for internal use -- don't use (returns APE_FILE_INFO *) [ignored, ignored]
End Enum

Private hAPE                    As APE_HANDLE

Private lngSamplerate           As Long
Private intChannels             As Integer
Private lngBitrate              As Long
Private intBitsPerSample        As Integer
Private lngLength               As Long
Private lngBlocks               As Long
Private lngBlockAlign           As Long
Private blnEOS                  As Boolean

Private clsTags                 As StreamTags

Private udtWFX                  As MMWAVEFORMATEX

' *************************************
' * constructor
' *************************************

Private Sub Class_Initialize()
    Set clsTags = New StreamTags
End Sub

Private Sub Class_Terminate()
    ISoundStream_CloseSource
End Sub

' *************************************
' * Stream properties
' *************************************

Private Property Get ISoundInfo_Bitrate() As Long
    ISoundInfo_Bitrate = lngBitrate * 1000
End Property

Private Property Get ISoundInfo_BitsPerSample() As Integer
    ISoundInfo_BitsPerSample = intBitsPerSample
End Property

Private Property Get ISoundInfo_Channels() As Integer
    ISoundInfo_Channels = intChannels
End Property

Private Property Get ISoundInfo_Duration() As Long
    ISoundInfo_Duration = lngLength
End Property

Private Property Get ISoundInfo_Force16Bit() As Boolean
    ISoundInfo_Force16Bit = False
End Property

Private Property Get ISoundInfo_Position() As Long
    Dim ret As Long

    ret = APEDecompress_GetInfo(hAPE, APE_DECOMPRESS_CURRENT_MS, 0, 0)
    ISoundInfo_Position = ret
End Property

Private Property Get ISoundInfo_Samplerate() As Long
    ISoundInfo_Samplerate = lngSamplerate
End Property

Private Property Get ISoundInfo_Tags() As StreamTags
    Set ISoundInfo_Tags = clsTags
End Property

Private Property Get ISoundStream_EndOfStream() As Boolean
    ISoundStream_EndOfStream = blnEOS
End Property

Private Property Get ISoundStream_Seekable() As Boolean
    ISoundStream_Seekable = True
End Property

Private Property Get ISoundStream_Extensions() As String()
    Dim strExt(0)   As String

    strExt(0) = "ape"
    ISoundStream_Extensions = strExt
End Property

Private Property Get ISoundStream_Info() As ISoundInfo
    Set ISoundStream_Info = Me
End Property

Private Property Get ISoundStream_WFXPointer() As Long
    ISoundStream_WFXPointer = VarPtr(udtWFX)
End Property

' *************************************
' * main functions
' *************************************

Private Function ISoundStream_CloseSource() As STREAM_ERR
    If hAPE <> 0 Then
        APEDecompress_Destroy hAPE
        hAPE = ape_invalid_handle
    End If

    ISoundStream_CloseSource = STREAM_OK
End Function

Private Function ISoundStream_OpenSource( _
    ByVal Source As String _
) As STREAM_ERR

    Dim lngErr  As Long

    If Not IsValidFile(Source) Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    If hAPE <> 0 Then ISoundStream_CloseSource

    hAPE = APEDecompress_Create(Source, lngErr)
    If lngErr <> 0 Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    intBitsPerSample = APEDecompress_GetInfo(hAPE, APE_INFO_BITS_PER_SAMPLE, 0, 0)
    lngBlockAlign = APEDecompress_GetInfo(hAPE, APE_INFO_BLOCK_ALIGN, 0, 0)
    lngSamplerate = APEDecompress_GetInfo(hAPE, APE_INFO_SAMPLE_RATE, 0, 0)
    intChannels = APEDecompress_GetInfo(hAPE, APE_INFO_CHANNELS, 0, 0)
    lngBitrate = APEDecompress_GetInfo(hAPE, APE_INFO_AVERAGE_BITRATE, 0, 0)
    lngLength = APEDecompress_GetInfo(hAPE, APE_INFO_LENGTH_MS, 0, 0)
    lngBlocks = APEDecompress_GetInfo(hAPE, APE_INFO_TOTAL_BLOCKS, 0, 0)

    ReadAPETags Source

    udtWFX = CreateWFX(lngSamplerate, intChannels, intBitsPerSample)

    blnEOS = False

    ISoundStream_OpenSource = STREAM_OK
End Function

Private Function ISoundStream_Read( _
    ByVal pBuffer As Long, _
    ByVal cBytes As Long, _
    cBytesRead As Long _
) As STREAM_ERR

    Dim ret             As Long
    Dim lngRetrieved    As Long

    ret = APEDecompress_GetData(hAPE, pBuffer, cBytes / udtWFX.nBlockAlign, lngRetrieved)

    If lngRetrieved = 0 Then
        ISoundStream_Read = STREAM_END_OF_STREAM
        cBytesRead = 0
        blnEOS = True
        Exit Function
    End If

    cBytesRead = lngRetrieved * udtWFX.nBlockAlign

    ISoundStream_Read = STREAM_OK
End Function

Private Function ISoundStream_SeekTo( _
    ByVal value As Long, _
    mode As SEEK_MODE _
) As STREAM_ERR

    Dim ret         As Long
    Dim lngBlockPos As Long

    Select Case mode

        Case SEEK_PERCENT
            If value < 0 Or value > 100 Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            lngBlockPos = lngBlocks / 100 * value

        Case SEEK_SECONDS
            If value < 0 Or value > (ISoundInfo_Duration / 1000) Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            lngBlockPos = lngBlocks / 100 * (value / (ISoundInfo_Duration / 1000) * 100)

    End Select

    ret = APEDecompress_Seek(hAPE, lngBlockPos)

    ISoundStream_SeekTo = STREAM_OK
End Function

' *************************************
' * helpers
' *************************************

Private Sub ReadAPETags( _
    strFile As String _
)

    Dim fh          As Integer: fh = FreeFile

    Dim udtHdr      As APE_TAG_FOOTER
    Dim udtItem     As APE_TAG_ITEM

    Dim btTags()    As Byte
    Dim intPos      As Integer
    Dim j           As Integer
    Dim i           As Integer

    Dim strItemName As String
    Dim strItemVal  As String
    Dim Bt          As Byte

    clsTags.ClearTags

    Open strFile For Binary As #fh

        Seek #fh, LOF(fh) - Len(udtHdr) + 1
        Get #fh, , udtHdr

        If udtHdr.preample = "APETAGEX" Then

            ReDim btTags(udtHdr.tagsize - 1) As Byte
            Seek #fh, LOF(fh) - udtHdr.tagsize + 1
            Get #fh, , btTags

        End If

    Close #fh

    If Not udtHdr.preample = "APETAGEX" Then
        Exit Sub
    End If

    For i = 1 To udtHdr.cTags

        strItemName = ""
        strItemVal = ""

        DXCopyMemory udtItem, btTags(intPos), Len(udtItem)
        ' skip flags
        intPos = intPos + 8

        Bt = 1
        Do While Not Bt = 0

            Bt = btTags(intPos)
            intPos = intPos + 1

            strItemName = strItemName & Chr$(Bt)

        Loop

        For j = intPos To intPos + udtItem.size - 1

            Bt = btTags(intPos)
            intPos = intPos + 1
            strItemVal = strItemVal & Chr$(Bt)

        Next

        clsTags.AddTag RemNullChars(strItemName), _
                       RemNullChars(strItemVal)

    Next
End Sub
