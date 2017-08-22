VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamOGG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Ogg Vorbis Decoder
'
' uses the high level OV API vorbisfile.dll,
' which relies on ogg.dll and vorbis.dll.
' Because the OV API expects either an C file handle
' or callbacks (cdecl convention), we simply
' use the MS C Runtime msvcrt.dll to obtain
' a C file handle.

Implements ISoundStream
Implements ISoundInfo

Private Type vorbis_info
    Version             As Long
    channels            As Long
    rate                As Long
    bitrate_upper       As Long
    bitrate_nominal     As Long
    bitrate_lower       As Long
    bitrate_window      As Long
    codec_setup         As Long
End Type

Private cCppRT          As clsDLL
Private cOgg            As clsDLL
Private cVorbis         As clsDLL
Private cVorbisFile     As clsDLL

Private blnReady        As Boolean

Private udtWFX          As MMWAVEFORMATEX

Private udtInfo         As vorbis_info
Private vf(64 * 1024&)  As Byte

Private lngDuration     As Long
Private lngCurPos       As Long

Private hFile           As Long
Private blnEOS          As Boolean

Private clsTags         As StreamTags

' ***********************************
' * constructor
' ***********************************

Private Sub Class_Initialize()
    Set cOgg = New clsDLL
    Set cCppRT = New clsDLL
    Set cVorbis = New clsDLL
    Set cVorbisFile = New clsDLL
    Set clsTags = New StreamTags

    If Not cCppRT.LoadDLL("msvcrt.dll") Then Exit Sub
    If Not cOgg.LoadDLL("ogg.dll") Then Exit Sub
    If Not cVorbis.LoadDLL("vorbis.dll") Then Exit Sub
    If Not cVorbisFile.LoadDLL("vorbisfile.dll") Then Exit Sub

    cOgg.IsCDECL = True
    cCppRT.IsCDECL = True
    cVorbis.IsCDECL = True
    cVorbisFile.IsCDECL = True

    blnReady = True
End Sub

Private Sub Class_Terminate()
    ISoundStream_CloseSource

    cVorbisFile.UnloadDLL
    cVorbis.UnloadDLL
    cOgg.UnloadDLL
    cCppRT.UnloadDLL
End Sub

' ***********************************
' * Stream properties
' ***********************************

Private Property Get ISoundInfo_Bitrate() As Long
    ISoundInfo_Bitrate = udtInfo.bitrate_nominal
End Property

Private Property Get ISoundInfo_BitsPerSample() As Integer
    ISoundInfo_BitsPerSample = 16
End Property

Private Property Get ISoundInfo_Channels() As Integer
    ISoundInfo_Channels = udtInfo.channels
End Property

Private Property Get ISoundInfo_Duration() As Long
    ISoundInfo_Duration = lngDuration
End Property

Private Property Get ISoundInfo_Force16Bit() As Boolean
    ISoundInfo_Force16Bit = True
End Property

Private Property Get ISoundInfo_Position() As Long
    ISoundInfo_Position = lngCurPos
End Property

Private Property Get ISoundInfo_Samplerate() As Long
    ISoundInfo_Samplerate = udtInfo.rate
End Property

Private Property Get ISoundInfo_Tags() As StreamTags
    Set ISoundInfo_Tags = clsTags
End Property

Private Property Get ISoundStream_EndOfStream() As Boolean
    ISoundStream_EndOfStream = blnEOS
End Property

Private Property Get ISoundStream_Extensions() As String()
    Dim strExt(0)   As String

    strExt(0) = "ogg"
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

' ***********************************
' * main functions
' ***********************************

Private Function ISoundStream_CloseSource() As STREAM_ERR
    If hFile <> 0 Then
        cVorbisFile.CallFunc "ov_clear", VarPtr(vf(0))
        cCppRT.CallFunc "fclose", hFile
        hFile = 0
    End If

    ISoundStream_CloseSource = STREAM_OK
End Function

Private Function ISoundStream_OpenSource( _
    ByVal Source As String _
) As STREAM_ERR

    Dim ret         As Long
    Dim btFile()    As Byte
    Dim btMode(2)   As Byte
    Dim strTag      As String

    If Not blnReady Then
        ISoundStream_OpenSource = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    If hFile <> 0 Then ISoundStream_CloseSource

    If Not IsValidFile(Source) Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    btFile = StrConv(Source & Chr$(0), vbFromUnicode)

    btMode(0) = Asc("r")    ' read access
    btMode(1) = Asc("b")    ' binary

    ' open the file through the C Runtime function
    ' "fopen", because VorbisFile only accepts these
    hFile = cCppRT.CallFunc("fopen", VarPtr(btFile(0)), VarPtr(btMode(0)))
    If hFile = 0 Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    ' valid Ogg Stream?
    If cVorbisFile.CallFunc("ov_open", hFile, VarPtr(vf(0)), 0, 0) < 0 Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    ' Info about the stream
    ret = cVorbisFile.CallFunc("ov_info", VarPtr(vf(0)), -1)
    If ret = 0 Then
        cVorbisFile.CallFunc "ov_clear", VarPtr(vf(0))
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    DXCopyMemory udtInfo, ByVal ret, Len(udtInfo)

    ' valid samplerate?
    If udtInfo.rate < 1 Then
        cVorbisFile.CallFunc "ov_clear", VarPtr(vf(0))
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    ' only up to 2 channels allowed
    If udtInfo.channels > 2 Then
        cVorbisFile.CallFunc "ov_clear", VarPtr(vf(0))
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    ' length of the file
    ret = cVorbisFile.CallFunc("ov_pcm_total", VarPtr(vf(0)), -1)
    lngDuration = ret / udtInfo.rate * 1000

    blnEOS = False
    lngCurPos = 0

    clsTags.ClearTags

    strTag = GetVorbisTag("artist")
    If strTag <> "" Then clsTags.AddTag "artist", strTag

    strTag = GetVorbisTag("album")
    If strTag <> "" Then clsTags.AddTag "album", strTag

    strTag = GetVorbisTag("title")
    If strTag <> "" Then clsTags.AddTag "title", strTag

    strTag = GetVorbisTag("genre")
    If strTag <> "" Then clsTags.AddTag "genre", strTag

    strTag = GetVorbisTag("tracknumber")
    If strTag <> "" Then clsTags.AddTag "tracknumber", strTag

    strTag = GetVorbisTag("description")
    If strTag <> "" Then clsTags.AddTag "description", strTag

    udtWFX = CreateWFX(udtInfo.rate, CInt(udtInfo.channels), 16)

    ISoundStream_OpenSource = STREAM_OK
End Function

Private Function ISoundStream_Read( _
    ByVal pBuffer As Long, _
    ByVal cBytes As Long, _
    cBytesRead As Long _
) As STREAM_ERR

    Dim ret             As Long
    Dim current_section As Long

    ' read PCM data
    ret = cVorbisFile.CallFunc("ov_read", _
                               VarPtr(vf(0)), _
                               pBuffer, _
                               cBytes, _
                               0, 2, 1, _
                               VarPtr(current_section))

    If ret = 0 Then
        blnEOS = True
        ISoundStream_Read = STREAM_END_OF_STREAM
    ElseIf ret > 0 Then
        lngCurPos = lngCurPos + ret * 1000 / udtInfo.rate / 2 / udtInfo.channels
        cBytesRead = ret
        ISoundStream_Read = STREAM_OK
    End If
End Function

Private Function ISoundStream_SeekTo( _
    ByVal value As Long, _
    mode As SEEK_MODE _
) As STREAM_ERR

    Dim dblTime     As Double
    Dim lngTimeHi   As Long
    Dim lngTimeLo   As Long
    Dim ret         As Long

    Select Case mode

        Case SEEK_PERCENT
            If value < 0 Or value > 100 Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            dblTime = (ISoundInfo_Duration / 1000) / 100 * value

        Case SEEK_SECONDS
            If value < 0 Or value > (ISoundInfo_Duration / 1000) Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            dblTime = value

    End Select

    lngCurPos = dblTime * 1000

    DXCopyMemory lngTimeLo, dblTime, 4
    DXCopyMemory lngTimeHi, ByVal VarPtr(dblTime) + 4, 4

    ret = cVorbisFile.CallFunc("ov_time_seek", VarPtr(vf(0)), lngTimeLo, lngTimeHi)

    blnEOS = False
    ISoundStream_SeekTo = STREAM_OK
End Function

' ***********************************
' * helpers
' ***********************************

Private Function GetVorbisTag( _
    ByVal strTag As String _
) As String

    Dim pComment    As Long
    Dim btTag()     As Byte
    Dim pTag        As Long

    pComment = cVorbisFile.CallFunc("ov_comment", VarPtr(vf(0)), -1)
    If pComment <> 0 Then
        btTag = StrConv(strTag & Chr$(0), vbFromUnicode)

        pTag = cVorbis.CallFunc("vorbis_comment_query", pComment, VarPtr(btTag(0)), 0)
        If pTag <> 0 Then
            GetVorbisTag = Ptr2Str(pTag)
        End If
    End If
End Function

Private Function Ptr2Str( _
    ByVal ptr As Long _
) As String

    Dim strBuffer   As String
    Dim btChr       As Byte
    Dim lngPos      As Long

    btChr = 1
    Do While btChr > 0
        DXCopyMemory btChr, ByVal ptr + lngPos, 1
        If btChr > 0 Then
            strBuffer = strBuffer & Chr$(btChr)
        End If
        lngPos = lngPos + 1
    Loop

    Ptr2Str = strBuffer
End Function
