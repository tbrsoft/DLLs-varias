VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderOGG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Ogg Vorbis Encoder
'
' translated from:
' http://open.nit.ca/wvstreams/docs/wvoggvorbis_8cc-source.html

Implements IEncoder

Public Enum OV_ENCMODE
    OV_ENC_ABR
    OV_ENC_QUALITY
End Enum

Private Type ogg_page
    pHeader                         As Long
    header_len                      As Long
    pBody                           As Long
    body_len                        As Long
End Type

Private clsOgg                      As clsDLL
Private clsVorbis                   As clsDLL

Private strArtist                   As String
Private strAlbum                    As String
Private strTitle                    As String

Private sngQuality                  As Single

Private lngBitrateNom               As Long
Private lngBitrateMin               As Long
Private lngBitrateMax               As Long

Private os                          As MEMORY
Private og                          As MEMORY
Private op                          As MEMORY
Private vc                          As MEMORY
Private vi                          As MEMORY
Private vd                          As MEMORY
Private vb                          As MEMORY

Private udeMode                     As OV_ENCMODE

Private hFOut                       As hFile

Private intChannels                 As Integer

Private blnOggReady                 As Boolean

Private Const SAMPLES_PER_ENC       As Long = 8192
Private Const STRUCT_SIZE           As Long = 128& * 1024&

' *****************************************
' * constructor
' *****************************************

Private Sub Class_Initialize()
    lngBitrateMax = -1
    lngBitrateMin = -1
    lngBitrateNom = 128000

    sngQuality = 0.4

    udeMode = OV_ENC_QUALITY

    Set clsOgg = New clsDLL
    Set clsVorbis = New clsDLL

    If Not LoadOggLibs Then
        UnloadOggLibs
    Else
        blnOggReady = True
    End If
End Sub

Private Sub Class_Terminate()
    UnloadOggLibs
End Sub

' *****************************************
' * Encoder settings
' *****************************************

Public Property Get EncoderMode() As OV_ENCMODE
    EncoderMode = udeMode
End Property

Public Property Let EncoderMode( _
    ByVal ude As OV_ENCMODE _
)

    udeMode = ude
End Property

Public Property Get BitrateMaximal() As Long
    BitrateMaximal = lngBitrateMax
End Property

Public Property Let BitrateMaximal( _
    ByVal lng As Long _
)

    lngBitrateMax = lng
End Property

Public Property Get BitrateMinimal() As Long
    BitrateMinimal = lngBitrateMin
End Property

Public Property Let BitrateMinimal( _
    ByVal lng As Long _
)

    lngBitrateMin = lng
End Property

Public Property Get BitrateNominal() As Long
    BitrateNominal = lngBitrateNom
End Property

Public Property Let BitrateNominal( _
    ByVal lng As Long _
)

    lngBitrateNom = lng
End Property

Public Property Get Quality() As Single
    Quality = sngQuality
End Property

Public Property Let Quality( _
    ByVal sng As Single _
)

    If sng < 0 Or sng > 1 Then Exit Property
    sngQuality = sng
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

Private Property Get IEncoder_Title() As String
    IEncoder_Title = strTitle
End Property

Private Property Let IEncoder_Title( _
    ByVal RHS As String _
)

    strTitle = RHS
End Property

Private Property Get IEncoder_Extension() As String
    IEncoder_Extension = "ogg"
End Property

' *****************************************
' * main functions
' *****************************************

Private Function IEncoder_Init( _
    ByVal dstFilename As String, _
    ByVal srcSampleRate As Long, _
    ByVal srcChannels As Integer, _
    ByVal srcBitsPerSample As Integer, _
    cbPCMData As Long _
) As STREAM_ERR

    Dim lngSerial                       As Long
    Dim lngRet                          As Long
    Dim btHeader(STRUCT_SIZE - 1)       As Byte
    Dim btHeaderComm(STRUCT_SIZE - 1)   As Byte
    Dim btHeaderCode(STRUCT_SIZE - 1)   As Byte

    IEncoder_Init = STREAM_NOT_SUPPORTED

    If Not blnOggReady Then
        Exit Function
    End If

    If srcChannels > 2 Or srcChannels < 1 Then _
        Exit Function
    If srcBitsPerSample < 16 Or srcBitsPerSample > 16 Then
        Exit Function
    End If

    ' *************************************
    ' * preparation
    ' *************************************
    AllocStructs

    lngSerial = Rnd() * 100000

    hFOut = FileOpen(dstFilename, GENERIC_WRITE, , CREATE_ALWAYS)
    If hFOut.handle = INVALID_HANDLE Then
        Exit Function
    End If

    ' *************************************
    ' * init Ogg Vorbis Encoder
    ' *************************************

    ' Vorbis Codec
    vorbis_info_init vi.pointer

    ' Vorbis Bitrate Management
    Select Case True

        Case udeMode = OV_ENC_QUALITY
            lngRet = vorbis_encode_init_vbr(vi.pointer, _
                                            srcChannels, _
                                            srcSampleRate, _
                                            sngQuality)

        Case udeMode = OV_ENC_ABR
            lngRet = vorbis_encode_init(vi.pointer, _
                                        srcChannels, _
                                        srcSampleRate, _
                                        lngBitrateMax, _
                                        lngBitrateNom, _
                                        lngBitrateMin)

        Case Else
            lngRet = -1

    End Select

    If lngRet < 0 Then GoTo OnError

    ' Vorbis Comments
    vorbis_comment_init vc.pointer
    vorbis_comment_add_tag vc.pointer, "artist", strArtist
    vorbis_comment_add_tag vc.pointer, "album", strAlbum
    vorbis_comment_add_tag vc.pointer, "title", strTitle

    ' Vorbis DSP
    lngRet = vorbis_analysis_init(vd.pointer, vi.pointer)
    If lngRet <> 0 Then GoTo OnError

    lngRet = vorbis_block_init(vd.pointer, vb.pointer)
    If lngRet <> 0 Then GoTo OnError

    ' Ogg Bitstream
    lngRet = ogg_stream_init(os.pointer, lngSerial)
    If lngRet <> 0 Then GoTo OnError

    ' Vorbis Header
    vorbis_analysis_headerout vd.pointer, _
                              vc.pointer, _
                              VarPtr(btHeader(0)), _
                              VarPtr(btHeaderComm(0)), _
                              VarPtr(btHeaderCode(0))

    ogg_stream_packetin os.pointer, VarPtr(btHeader(0))
    ogg_stream_packetin os.pointer, VarPtr(btHeaderComm(0))
    ogg_stream_packetin os.pointer, VarPtr(btHeaderCode(0))

    ' flush headers, so following data will
    ' start on a new page
    If Not WriteStream(True) Then
        GoTo OnError
    End If

    ' 1024 16-Bit Samples per encode
    cbPCMData = SAMPLES_PER_ENC * 2

    intChannels = srcChannels

    IEncoder_Init = STREAM_OK
    Exit Function

OnError:
    FreeStructs
End Function

Private Function IEncoder_Encode( _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As STREAM_ERR

    Dim intSamples()    As Integer
    Dim sngSamples()    As Single
    Dim cSamples        As Integer
    '
    Dim ppBuffers       As Long
    Dim pBuffers()      As Long
    '
    Dim j               As Long
    Dim i               As Long

    ' don't use more data then we've defined
    ' at encoder init
    If bytes > SAMPLES_PER_ENC * 2 Then
        bytes = SAMPLES_PER_ENC * 2
    End If

    cSamples = bytes / 2

    ReDim intSamples(cSamples - 1) As Integer
    ReDim sngSamples(cSamples - 1) As Single

    ' get a buffer for the data.
    ' returned value by vorbis_analysis_buffer is
    ' an array of pointers to buffers for each
    ' channel (stereo = 2).
    ppBuffers = vorbis_analysis_buffer(vd.pointer, cSamples / intChannels)
    If ppBuffers = 0 Then
        IEncoder_Encode = STREAM_CANCELED
        Exit Function
    End If

    ' Channel Buffer Array
    ReDim pBuffers(intChannels - 1) As Long
    DXCopyMemory pBuffers(0), ByVal ppBuffers, intChannels * 4

    ' make PCM data useable
    DXCopyMemory intSamples(0), ByVal ptr, bytes

    ' Vorbis wants the channel data in seperate buffers,
    ' so we need to deinterleave them
    ' (LRLRLR => LLL RRR)
    ' on top of that Vorbis wants floats [-1,1]...
    If intChannels = 2 Then
        ' Channel L
        j = 0
        For i = 0 To cSamples / 2 - 1
            sngSamples(i) = intSamples(j) / 32768#
            j = j + 2
        Next

        ' Channel R
        j = 1
        For i = cSamples / 2 To cSamples - 1
            sngSamples(i) = intSamples(j) / 32768#
            j = j + 2
        Next
    Else
        ' Mono, no deinterleaving
        For i = 0 To cSamples - 1
            sngSamples(i) = intSamples(i) / 32768#
        Next
    End If

    If intChannels = 2 Then
        DXCopyMemory ByVal pBuffers(0), sngSamples(0), cSamples / 2 * 4
        DXCopyMemory ByVal pBuffers(1), sngSamples(cSamples / 2), cSamples / 2 * 4
    Else
        DXCopyMemory ByVal pBuffers(0), sngSamples(0), cSamples * 4
    End If

    ' tell Vorbis how many samples we've written per channel
    vorbis_analysis_wrote vd.pointer, cSamples / intChannels

    ' encode samples and write pages to the file
    ProcessAudio
End Function

Private Function IEncoder_DeInit() As STREAM_ERR
    ' tell Vorbis we're at the end of the stream
    vorbis_analysis_wrote vd.pointer, 0

    ' flush pages
    ProcessAudio

    ogg_stream_clear os.pointer
    vorbis_block_clear vb.pointer
    vorbis_dsp_clear vd.pointer
    vorbis_comment_clear vc.pointer
    vorbis_info_clear vi.pointer

    FreeStructs
    FileClose hFOut

    IEncoder_DeInit = STREAM_OK
End Function

' *****************************************
' * OGG/Vorbis Stream Output
' *****************************************

Private Function WriteStream( _
    ByVal Flush As Boolean _
) As Boolean

    Dim udtPage     As ogg_page
    Dim lngRet      As Long

    Do
        If Flush Then
            lngRet = ogg_stream_flush(os.pointer, udtPage)
        Else
            lngRet = ogg_stream_pageout(os.pointer, udtPage)
        End If

        If lngRet = 0 Then
            Exit Do
        ElseIf lngRet < 0 Then
            Exit Function
        End If

        FileWrite hFOut, udtPage.pHeader, udtPage.header_len
        FileWrite hFOut, udtPage.pBody, udtPage.body_len
    Loop

    WriteStream = True
End Function

Private Function ProcessAudio() As Boolean
    Dim lngRet      As Long

    Do While vorbis_analysis_blockout(vd.pointer, vb.pointer) = 1
        lngRet = vorbis_analysis(vb.pointer, 0)
        If lngRet < 0 Then Exit Function

        lngRet = vorbis_bitrate_addblock(vb.pointer)
        If lngRet < 0 Then Exit Function

        Do While vorbis_bitrate_flushpacket(vd.pointer, op.pointer) > 0
            ogg_stream_packetin os.pointer, op.pointer
            If Not WriteStream(False) Then
                Exit Function
            End If
        Loop
    Loop

    ProcessAudio = True
End Function

' *****************************************
' * OGG/Vorbis API Wrapper
' *****************************************

Private Sub vorbis_info_init( _
    ByVal VorbisInfo As Long _
)

    clsVorbis.CallFunc "vorbis_info_init", VorbisInfo
End Sub

Private Sub vorbis_comment_init( _
    ByVal VorbisComment As Long _
)

    clsVorbis.CallFunc "vorbis_comment_init", VorbisComment
End Sub

Private Function vorbis_encode_init_vbr( _
    ByVal VorbisInfo As Long, _
    ByVal channels As Long, _
    ByVal samplerate As Long, _
    ByVal base_quality As Single _
) As Long

    Dim lngQuality  As Long
    DXCopyMemory lngQuality, base_quality, 4

    vorbis_encode_init_vbr = clsVorbis.CallFunc("vorbis_encode_init_vbr", _
                                                VorbisInfo, _
                                                channels, _
                                                samplerate, _
                                                lngQuality)
End Function

Private Function vorbis_encode_init( _
    ByVal VorbisInfo As Long, _
    ByVal channes As Long, _
    ByVal samplerate As Long, _
    ByVal max_bitrate As Long, _
    ByVal nom_bitrate As Long, _
    ByVal min_bitrate As Long _
) As Long

    vorbis_encode_init = clsVorbis.CallFunc("vorbis_encode_init", _
                                            VorbisInfo, _
                                            channes, _
                                            samplerate, _
                                            max_bitrate, _
                                            nom_bitrate, _
                                            min_bitrate)
End Function

Private Function vorbis_analysis_init( _
    ByVal VorbisDSP As Long, _
    ByVal VorbisInfo As Long _
) As Long

    vorbis_analysis_init = clsVorbis.CallFunc("vorbis_analysis_init", _
                                              VorbisDSP, _
                                              VorbisInfo)
End Function

Private Function vorbis_block_init( _
    ByVal VorbisDSP As Long, _
    ByVal VorbisBlock As Long _
) As Long

    vorbis_block_init = clsVorbis.CallFunc("vorbis_block_init", _
                                           VorbisDSP, _
                                           VorbisBlock)
End Function

Private Sub vorbis_block_clear( _
    ByVal VorbisBlock As Long _
)
    clsVorbis.CallFunc "vorbis_block_clear", VorbisBlock
End Sub

Private Sub vorbis_dsp_clear( _
    ByVal VorbisDSP As Long _
)

    clsVorbis.CallFunc "vorbis_dsp_clear", VorbisDSP
End Sub

Private Sub vorbis_comment_clear( _
    ByVal VorbisComment As Long _
)

    clsVorbis.CallFunc "vorbis_comment_clear", VorbisComment
End Sub

Private Sub vorbis_info_clear( _
    ByVal VorbisInfo As Long _
)

    clsVorbis.CallFunc "vorbis_info_clear", VorbisInfo
End Sub

Private Sub vorbis_comment_add_tag( _
    ByVal VorbisComment As Long, _
    ByVal Tag As String, _
    ByVal TagValue As String _
)

    Dim btTag()     As Byte
    Dim btTagVal()  As Byte

    btTag = StrConv(Tag & Chr$(0), vbFromUnicode)
    btTagVal = StrConv(TagValue & Chr$(0), vbFromUnicode)

    clsVorbis.CallFunc "vorbis_comment_add_tag", _
                       VorbisComment, _
                       VarPtr(btTag(0)), _
                       VarPtr(btTagVal(0))
End Sub

Private Function vorbis_analysis_buffer( _
    ByVal VorbisDSP As Long, _
    ByVal VorbisSamples As Long _
) As Long

    vorbis_analysis_buffer = clsVorbis.CallFunc("vorbis_analysis_buffer", _
                                                VorbisDSP, _
                                                VorbisSamples)
End Function

Private Sub vorbis_analysis_wrote( _
    ByVal VorbisDSP As Long, _
    ByVal VorbisSamples As Long _
)

    clsVorbis.CallFunc "vorbis_analysis_wrote", _
                       VorbisDSP, _
                       VorbisSamples
End Sub

Private Function vorbis_analysis_headerout( _
    ByVal VorbisDSP As Long, _
    ByVal VorbisComment As Long, _
    ByVal header As Long, _
    ByVal header_comm As Long, _
    ByVal header_code As Long _
) As Long

    vorbis_analysis_headerout = clsVorbis.CallFunc("vorbis_analysis_headerout", _
                                                   VorbisDSP, _
                                                   VorbisComment, _
                                                   header, _
                                                   header_comm, _
                                                   header_code)
End Function

Private Function vorbis_analysis_blockout( _
    ByVal VorbisDSP As Long, _
    ByVal VorbisBlock As Long _
) As Long

    vorbis_analysis_blockout = clsVorbis.CallFunc("vorbis_analysis_blockout", _
                                                  VorbisDSP, _
                                                  VorbisBlock)
End Function

Private Function vorbis_analysis( _
    ByVal VorbisBlock As Long, _
    ByVal packet As Long _
) As Long

    vorbis_analysis = clsVorbis.CallFunc("vorbis_analysis", _
                                         VorbisBlock, _
                                         packet)
End Function

Private Function vorbis_bitrate_addblock( _
    ByVal VorbisBlock As Long _
) As Long

    vorbis_bitrate_addblock = clsVorbis.CallFunc("vorbis_bitrate_addblock", _
                                                 VorbisBlock)
End Function

Private Function vorbis_bitrate_flushpacket( _
    ByVal VorbisDSP As Long, _
    ByVal packet As Long _
) As Long

    vorbis_bitrate_flushpacket = clsVorbis.CallFunc("vorbis_bitrate_flushpacket", _
                                                    VorbisDSP, _
                                                    packet)
End Function

Private Function ogg_stream_init( _
    ByVal OggStream As Long, _
    ByVal SerialNo As Long _
) As Long

    ogg_stream_init = clsOgg.CallFunc("ogg_stream_init", _
                                      OggStream, _
                                      SerialNo)
End Function

Private Sub ogg_stream_clear( _
    ByVal OggStream As Long _
)

    clsOgg.CallFunc "ogg_stream_clear", OggStream
End Sub

Private Function ogg_stream_packetin( _
    ByVal OggStream As Long, _
    ByVal packet As Long _
) As Long

    ogg_stream_packetin = clsOgg.CallFunc("ogg_stream_packetin", _
                                          OggStream, _
                                          packet)
End Function

Private Function ogg_stream_flush( _
    ByVal OggStream As Long, _
    page As ogg_page _
) As Long

    ogg_stream_flush = clsOgg.CallFunc("ogg_stream_flush", _
                                       OggStream, _
                                       VarPtr(page))
End Function

Private Function ogg_stream_pageout( _
    ByVal OggStream As Long, _
    page As ogg_page _
) As Long

    ogg_stream_pageout = clsOgg.CallFunc("ogg_stream_pageout", _
                                         OggStream, _
                                         VarPtr(page))
End Function

' *****************************************
' * helpers
' *****************************************

Private Function LoadOggLibs() As Boolean
    If Not clsOgg.LoadDLL("ogg.dll") Then Exit Function
    If Not clsVorbis.LoadDLL("vorbis.dll") Then Exit Function

    clsOgg.IsCDECL = True
    clsVorbis.IsCDECL = True

    LoadOggLibs = True
End Function

Private Sub UnloadOggLibs()
    clsOgg.UnloadDLL
    clsVorbis.UnloadDLL
End Sub

Private Sub AllocStructs()
    os = AllocateMem(STRUCT_SIZE)
    og = AllocateMem(STRUCT_SIZE)
    op = AllocateMem(STRUCT_SIZE)
    vc = AllocateMem(STRUCT_SIZE)
    vi = AllocateMem(STRUCT_SIZE)
    vd = AllocateMem(STRUCT_SIZE)
    vb = AllocateMem(STRUCT_SIZE)
End Sub

Private Sub FreeStructs()
    FreeMem os
    FreeMem og
    FreeMem op
    FreeMem vc
    FreeMem vi
    FreeMem vd
    FreeMem vb
End Sub
