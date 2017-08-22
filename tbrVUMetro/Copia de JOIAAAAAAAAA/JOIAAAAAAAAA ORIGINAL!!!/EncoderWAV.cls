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
Attribute VB_Exposed = False
Option Explicit

' WAV Encoder with ACM support for compressed formats
' (from my project "Monoton")

' this class implemented "IEncoder".
' to hold the size of the project down, I removed it.
' now the members of IEncoder have the prefix "encoder".


Private Const Extension         As String = "wav"
Private Const Description       As String = "Wave ACM Encoder"


Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal dwLen As Long _
)

Private Declare Sub ZeroMem Lib "kernel32" _
Alias "RtlZeroMemory" ( _
    pDst As Any, _
    ByVal dwLen As Long _
)

Private Declare Function IsBadReadPtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Private Declare Function IsBadWritePtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Private Declare Function CreateFile Lib "kernel32.dll" _
Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal HFILE As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any _
) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal HFILE As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Any _
) As Long

Private Declare Function SetFilePointer Lib "kernel32" ( _
    ByVal HFILE As Long, _
    ByVal lDistanceToMove As Long, _
    ByVal lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long _
) As Long

Private Declare Function GetFileSize Lib "kernel32" ( _
    ByVal HFILE As Long, _
    ByVal lpFileSizeHigh As Long _
) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long _
) As Long

Private Declare Function acmStreamPrepareHeader Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwPrepare As Long _
) As Long

Private Declare Function acmStreamUnprepareHeader Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwUnprepare As Long _
) As Long

Private Declare Function acmStreamOpen Lib "msacm32" ( _
    phas As Long, _
    ByVal had As Long, _
    pwfxSrc As Any, _
    pwfxDst As Any, _
    ByVal pwfltr As Long, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal fdwOpen As Long _
) As Long

Private Declare Function acmStreamSize Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal cbInput As Long, _
    pdwOutputBytes As Long, _
    ByVal fdwSize As Long _
) As Long

Private Declare Function acmStreamConvert Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwConvert As Long _
) As Long

Private Declare Function acmStreamReset Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal fdwReset As Long _
) As Long

Private Declare Function acmStreamClose Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal fdwClose As Long _
) As Long

Private Declare Function acmFormatChoose Lib "msacm32" _
Alias "acmFormatChooseA" ( _
    pfmtc As ACMFORMATCHOOSEA _
) As Long

Public Enum SND_RESULT
    SND_ERR_SUCCESS = 0
    SND_ERR_INVALID_SOURCE
    SND_ERR_INVALID_OUTPUT
    SND_ERR_INTERNAL
    SND_ERR_OUT_OF_RANGE
    SND_ERR_END_OF_STREAM
    SND_ERR_INVALID_TAG
    SND_ERR_INVALID_PARAM
    SND_ERR_TOO_BIG
    SND_ERR_NEED_MORE
    SND_ERR_WRITE_ERROR
    SND_ERR_UNKNOWN
End Enum

Public Enum ENCODER_METRICS
    ENC_M_PCM_PER_DELIVERY_MIN = 0
    ENC_M_PCM_PER_DELIVERY_MAX
    ENC_M_SUPPORT_TAGS
End Enum

Private Enum FILE_OPEN_METHOD
    CREATE_NEW = 1
    CREATE_ALWAYS = 2
    OPEN_EXISTING = 3
    OPEN_ALWAYS = 4
End Enum

Private Enum FILE_SHARE_RIGHTS
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
End Enum

Private Enum FILE_ACCESS_RIGHTS
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
End Enum

Private Enum SEEK_METHOD
    FILE_BEGIN = 0
    FILE_CURRENT = 1
    FILE_END = 2
End Enum

Private Enum HACMSTREAM
    INVALID_STREAM_HANDLE = 0
End Enum

Private Enum ACM_STREAMSIZEF
    ACM_STREAMSIZEF_DESTINATION = &H1
    ACM_STREAMSIZEF_SOURCE = &H0
    ACM_STREAMSIZEF_QUERYMASK = &HF
End Enum

Private Enum ACM_STREAMCONVERTF
    ACM_STREAMCONVERTF_BLOCKALIGN = &H4
    ACM_STREAMCONVERTF_START = &H10
    ACM_STREAMCONVERTF_END = &H20
End Enum

Private Const INVALID_HANDLE                As Long = -1

Private Const ACMFMTDET_FORMAT_CHARS        As Long = 128
Private Const ACMFMTTAGDET_FORMATTAG_CHARS  As Long = 48

Private Const ACM_FORMATENUMF_CONVERT       As Long = &H100000

Private Const WAVE_FORMAT_PCM               As Long = 1

' some codecs (e.g. Vorbis ACM) have very big
' WFX structs. so make the WFX buffer extra large.
Private Const WFXSize                       As Long = 128& * 1024&

Private Const WAV_RIFF                      As Long = 1179011410
Private Const WAV_WAVE                      As Long = 1163280727
Private Const WAV_DATA                      As Long = 1635017060
Private Const WAV_FMT                       As Long = 544501094

Private Type HFILE
    handle                                  As Long
    path                                    As String
End Type

Private Type ACMSTREAMHEADER
    cbStruct                                As Long
    fdwStatus                               As Long
    dwUser                                  As Long
    pbSrc                                   As Long
    cbSrcLength                             As Long
    cbSrcLengthUsed                         As Long
    dwSrcUser                               As Long
    pbDst                                   As Long
    cbDstLength                             As Long
    cbDstLengthUsed                         As Long
    dwDstUser                               As Long
    dwReservedDriver(9)                     As Long
End Type

Private Type ACMFORMATCHOOSEA
    cbStruct                                As Long
    fdwStyle                                As Long
    hwndOwner                               As Long
    pwfx                                    As Long
    cbwfx                                   As Long
    pszTitle                                As Long
    szFormatTag                             As String * ACMFMTTAGDET_FORMATTAG_CHARS
    szFormat                                As String * ACMFMTDET_FORMAT_CHARS
    pszName                                 As Long
    cchName                                 As Long
    fdwEnum                                 As Long
    pwfxEnum                                As Long
    hInstance                               As Long
    pszTemplateName                         As Long
    lCustData                               As Long
    pfnHook                                 As Long
    btSpace(1023)                           As Byte ' had some strange crashes...
End Type

Private Type MMWAVEFORMATEX
    wFormatTag                              As Integer
    nChannels                               As Integer
    nSamplesPerSec                          As Long
    nAvgBytesPerSec                         As Long
    nBlockAlign                             As Integer
    wBitsPerSample                          As Integer
    cbSize                                  As Integer
End Type

Private Type WAVRIFF
    RIFF                                    As Long         ' "RIFF"
    hdrlen                                  As Long
    WAVE                                    As Long         ' "WAVE"
    fmt                                     As Long         ' "fmt "
    chnksize                                As Long
End Type

Private udtWFXIn                            As MMWAVEFORMATEX
Private btWfxOut()                          As Byte
Private lngWFXOutLen                        As Long
Private strFormatID                         As String
Private strFormatTag                        As String

Private btInp()                             As Byte
Private btOut()                             As Byte
Private lngInpLen                           As Long
Private lngOutLen                           As Long

Private udtStreamHdr                        As ACMSTREAMHEADER
Private hStream                             As HACMSTREAM

Private hFOut                               As HFILE

Private blnWriteHeader                      As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    Dim wfx As MMWAVEFORMATEX

    ReDim btWfxOut(WFXSize - 1) As Byte
    blnWriteHeader = True

    wfx = CreateWFX(44100, 2, 16)

    SetFormat VarPtr(wfx), _
              Len(wfx), _
              "44.1 kHz 16 Bit Stereo", "PCM"

    hFOut.handle = INVALID_HANDLE
End Sub

Private Sub Class_Terminate()
    Encoder_EncoderClose
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

' write RIFF/WAVE/FMT /DATA chunks
Public Property Get WriteHeader( _
) As Boolean

    WriteHeader = blnWriteHeader
End Property

Public Property Let WriteHeader( _
ByVal bln As Boolean _
)

    If hFOut.handle = INVALID_HANDLE Then
        blnWriteHeader = bln
    Else
        Err.Raise 90, , "Encoder already running!"
    End If
End Property

Public Property Get FormatTag( _
) As String

    FormatTag = strFormatTag
End Property

Public Property Get FormatID( _
) As String

    FormatID = strFormatID
End Property

Public Sub SetPCMFormat( _
    ByVal samplerate As Long, _
    ByVal Channels As Integer, _
    Optional ByVal FormatID As String, _
    Optional ByVal FormatTag As String _
)

    Dim wfx As MMWAVEFORMATEX

    wfx = CreateWFX(samplerate, Channels, 16)

    If FormatID = "" Then
        With wfx
            strFormatID = Round(.nSamplesPerSec / 1000, 1) & " " & _
                          .wBitsPerSample & " Bit " & _
                          IIf(.nChannels = 2, "stereo", "mono")
        End With
    End If

    If FormatTag = "" Then
        FormatTag = "PCM"
    End If

    SetFormat VarPtr(wfx), Len(wfx), _
              FormatID, FormatTag
End Sub

Public Sub SetFormat( _
    ByVal wfx_ptr As Long, _
    ByVal wfx_len As Long, _
    Optional FormatID As String, _
    Optional FormatTag As String _
)

    If 0 = IsBadReadPtr(ByVal wfx_ptr, wfx_len) Then
        If wfx_len > UBound(btWfxOut) + 1 Then
            ReDim btWfxOut(wfx_len - 1) As Byte
        End If
        CpyMem btWfxOut(0), ByVal wfx_ptr, wfx_len

        lngWFXOutLen = wfx_len

        strFormatID = FormatID
        strFormatTag = FormatTag
    End If
End Sub

Public Function SelectFormat( _
    ByVal samplerate As Long, _
    ByVal Channels As Integer, _
    Optional ByVal hwndOwner As Long, _
    Optional ByVal Title As String = "WAV Format" _
) As SND_RESULT

    Dim btTitle()   As Byte
    Dim btWFX()     As Byte
    Dim wfx         As MMWAVEFORMATEX
    Dim wfxinp      As MMWAVEFORMATEX
    Dim udtFmtCh    As ACMFORMATCHOOSEA

    ReDim btWFX(WFXSize - 1) As Byte

    btTitle = StrConv(Title & Chr$(0), vbFromUnicode)

    With wfxinp
        .nSamplesPerSec = samplerate
        .nChannels = Channels
        .wBitsPerSample = 16
        .wFormatTag = WAVE_FORMAT_PCM
        .nBlockAlign = .nChannels * (.wBitsPerSample / 8)
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
    End With

    CpyMem btWFX(0), wfxinp, Len(wfxinp)

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
        CpyMem wfx, btWFX(0), Len(wfx)
        lngWFXOutLen = CLng(18) + CLng(wfx.cbSize)

        strFormatID = RemNullChars(udtFmtCh.szFormat)
        strFormatTag = RemNullChars(udtFmtCh.szFormatTag)

        SelectFormat = SND_ERR_SUCCESS
    Else
        SelectFormat = SND_ERR_UNKNOWN
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Encoder_Description( _
) As String

    Encoder_Description = Description
End Property

Public Property Get Encoder_Extension( _
) As String

    Encoder_Extension = Extension
End Property

Public Function Encoder_EncoderClose( _
) As SND_RESULT

    If hFOut.handle = INVALID_HANDLE Then
        Encoder_EncoderClose = SND_ERR_INVALID_OUTPUT
    Else
        ' ACM Flush
        udtStreamHdr.cbSrcLength = 0
        acmStreamConvert hStream, udtStreamHdr, ACM_STREAMCONVERTF_END
        udtStreamHdr.cbSrcLength = lngInpLen

        FileWrite hFOut, udtStreamHdr.pbDst, udtStreamHdr.cbDstLengthUsed
        If blnWriteHeader Then WriteWAVHdr

        acmStreamUnprepareHeader hStream, udtStreamHdr, 0
        acmStreamClose hStream, 0
        hStream = INVALID_STREAM_HANDLE

        FileClose hFOut

        Encoder_EncoderClose = SND_ERR_SUCCESS
    End If
End Function

Public Function Encoder_Encode( _
    ByVal buffer_ptr As Long, _
    ByVal buffer_len As Long, _
    buffer_used As Long _
) As SND_RESULT

    Dim lngRet      As Long
    Dim lngWritten  As Long

    If buffer_len > lngInpLen Then
        Encoder_Encode = SND_ERR_TOO_BIG
        Exit Function
    End If

    If 0 = IsBadReadPtr(ByVal buffer_ptr, buffer_len) Then
        CpyMem btInp(0), ByVal buffer_ptr, buffer_len
    End If

    udtStreamHdr.cbSrcLength = buffer_len

    lngRet = acmStreamConvert(hStream, _
                              udtStreamHdr, _
                              ACM_STREAMCONVERTF_BLOCKALIGN)

    If lngRet <> 0 Then
        If udtStreamHdr.cbSrcLengthUsed = 0 Then
            ' you don't have to handle that
            Encoder_Encode = SND_ERR_NEED_MORE
        Else
            Encoder_Encode = SND_ERR_SUCCESS
        End If
    Else
        lngWritten = FileWrite(hFOut, _
                               udtStreamHdr.pbDst, _
                               udtStreamHdr.cbDstLengthUsed)

        If lngWritten = udtStreamHdr.cbDstLengthUsed Then
            Encoder_Encode = SND_ERR_SUCCESS
        Else
            Encoder_Encode = SND_ERR_WRITE_ERROR
        End If
    End If

    buffer_used = udtStreamHdr.cbSrcLengthUsed
End Function

Public Function Encoder_EncoderMetrics( _
    ByVal info As ENCODER_METRICS _
) As Long

    Select Case True

        Case info = ENC_M_PCM_PER_DELIVERY_MAX
            If hFOut.handle = INVALID_HANDLE Then
                Encoder_EncoderMetrics = -1
            Else
                Encoder_EncoderMetrics = udtWFXIn.nAvgBytesPerSec
            End If

        Case info = ENC_M_PCM_PER_DELIVERY_MIN
            If hFOut.handle = INVALID_HANDLE Then
                Encoder_EncoderMetrics = -1
            Else
                Encoder_EncoderMetrics = udtWFXIn.nBlockAlign
            End If

        Case info = ENC_M_SUPPORT_TAGS

    End Select
End Function

Public Function Encoder_EncoderInit( _
    ByVal samplerate As Long, _
    ByVal Channels As Integer, _
    ByVal output As String _
) As SND_RESULT

    Dim lngRet  As Long
    Dim btHdr() As Byte

    ' vorsichtshalber
    Encoder_EncoderClose

    If samplerate < 1 Or Channels < 1 Then
        Encoder_EncoderInit = SND_ERR_INVALID_PARAM
        Exit Function
    End If

    udtWFXIn = CreateWFX(samplerate, Channels, 16)

    ' start ACM conversion
    lngRet = acmStreamOpen(hStream, 0, _
                           udtWFXIn, _
                           btWfxOut(0), _
                           0, 0, 0, 0)

    If lngRet <> 0 Then
        Encoder_EncoderInit = SND_ERR_INVALID_PARAM
        Exit Function
    End If

    lngInpLen = udtWFXIn.nAvgBytesPerSec

    ' needed input buffer size for
    ' wanted output buffer size
    lngRet = acmStreamSize(hStream, _
                           lngInpLen, _
                           lngOutLen, _
                           ACM_STREAMSIZEF_SOURCE)

    If lngRet <> 0 Then
        Encoder_EncoderInit = SND_ERR_INVALID_PARAM
        acmStreamClose hStream, 0
        Exit Function
    End If

    ReDim btInp(lngInpLen - 1) As Byte
    ReDim btOut(lngOutLen - 1) As Byte

    ' ACM stream header saves a pointer to
    ' the buffer, a stream needs to have
    ' at least 1 stream header.
    With udtStreamHdr
        .cbStruct = Len(udtStreamHdr)
        .cbDstLength = lngOutLen
        .cbSrcLength = lngInpLen
        .pbDst = VarPtr(btOut(0))
        .pbSrc = VarPtr(btInp(0))
    End With

    lngRet = acmStreamPrepareHeader(hStream, _
                                    udtStreamHdr, _
                                    0)

    If lngRet <> 0 Then
        Encoder_EncoderInit = SND_ERR_INVALID_PARAM
        acmStreamClose hStream, 0
        Exit Function
    End If

    hFOut = FileOpen(output, _
                     GENERIC_WRITE, , _
                     CREATE_ALWAYS)

    If hFOut.handle = INVALID_HANDLE Then
        Encoder_EncoderInit = SND_ERR_INVALID_OUTPUT
        Exit Function
    End If

    If blnWriteHeader Then
        ' space for the WAV header (we can't write
        ' the chunks yet because we don't know the
        ' final size of the file)
        ReDim btHdr(20 + lngWFXOutLen + 4 * 2 - 1) As Byte
        FileWrite hFOut, VarPtr(btHdr(0)), UBound(btHdr) + 1
    End If

    Encoder_EncoderInit = SND_ERR_SUCCESS
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Private Function CreateWFX( _
    sr As Long, _
    chs As Integer, _
    bps As Integer _
) As MMWAVEFORMATEX

    With CreateWFX
        .wFormatTag = WAVE_FORMAT_PCM
        .nChannels = chs
        .nSamplesPerSec = sr
        .wBitsPerSample = bps
        .nBlockAlign = chs * (bps / 8)
        .nAvgBytesPerSec = sr * .nBlockAlign
    End With
End Function

Private Function RemNullChars( _
    ByVal strString As String _
) As String

    If InStr(strString, Chr$(0)) > 0 Then
        RemNullChars = Left$(strString, _
                             InStr(strString, Chr$(0)) - 1)
    Else
        RemNullChars = strString
    End If
End Function

' The header of a WAV file is actually 4 bytes big
' ("RIFF"). ' They're made of chunks which can be
' placed all over the file, but mostly they're all
' in front of the data.
Private Sub WriteWAVHdr()
    Dim udtHdr      As WAVRIFF
    Dim lngDataID   As Long
    Dim lngData     As Long

    With udtHdr
        .RIFF = WAV_RIFF
        .WAVE = WAV_WAVE
        .fmt = WAV_FMT
        .hdrlen = FileLength(hFOut) - 8
        .chnksize = lngWFXOutLen
    End With

    FileSeek hFOut, 0, FILE_BEGIN
    FileWrite hFOut, VarPtr(udtHdr), Len(udtHdr)

    FileWrite hFOut, VarPtr(btWfxOut(0)), lngWFXOutLen

    lngDataID = WAV_DATA
    FileWrite hFOut, VarPtr(lngDataID), 4

    lngData = FileLength(hFOut) - (FilePosition(hFOut) + 4)
    FileWrite hFOut, VarPtr(lngData), 4
End Sub

' //////////////////////////////////////
' ////// API File Handling
' //////////////////////////////////////

Private Function IsValidFile( _
    ByVal strFile As String _
) As Boolean

    Dim hInp    As HFILE

    hInp = FileOpen(strFile, GENERIC_READ, FILE_SHARE_READ)
    IsValidFile = hInp.handle <> INVALID_HANDLE
    FileClose hInp
End Function

Private Function FileOpen( _
    ByVal strFile As String, _
    Optional access As FILE_ACCESS_RIGHTS = GENERIC_READ Or GENERIC_WRITE, _
    Optional share As FILE_SHARE_RIGHTS = FILE_SHARE_READ Or FILE_SHARE_WRITE, _
    Optional method As FILE_OPEN_METHOD = OPEN_EXISTING _
) As HFILE

    FileOpen.handle = CreateFile(strFile, _
                                 access, _
                                 share, _
                                 ByVal 0&, _
                                 method, _
                                 0, 0)

    FileOpen.path = strFile
End Function

Private Sub FileClose( _
    filehandle As HFILE _
)

    CloseHandle filehandle.handle
    filehandle.handle = INVALID_HANDLE
    filehandle.path = vbNullString
End Sub

Private Function FileRead( _
    filehandle As HFILE, _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As Long

    Dim dwRead  As Long
    Dim lngRet  As Long

    If filehandle.handle = INVALID_HANDLE Then
        Exit Function
    End If

    lngRet = ReadFile(filehandle.handle, _
                      ByVal ptr, _
                      bytes, _
                      dwRead, _
                      0&)

    If lngRet = 1 Then
        FileRead = dwRead
    Else
        FileRead = -1
    End If
End Function

Private Function FileWrite( _
    filehandle As HFILE, _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As Long

    Dim dwWritten   As Long
    Dim lngRet      As Long

    If filehandle.handle = INVALID_HANDLE Then
        Exit Function
    End If

    lngRet = WriteFile(filehandle.handle, _
                       ByVal ptr, _
                       bytes, _
                       dwWritten, _
                       0&)

    If lngRet = 1 Then
        FileWrite = dwWritten
    Else
        FileWrite = -1
    End If
End Function

Private Function FileSeek( _
    filehandle As HFILE, _
    ByVal bytes As Long, _
    ByVal method As SEEK_METHOD _
) As Long

    FileSeek = SetFilePointer(filehandle.handle, _
                              bytes, _
                              0, _
                              method)
End Function

Private Function FilePosition( _
    filehandle As HFILE _
) As Long

    FilePosition = FileSeek(filehandle, _
                            0, _
                            FILE_CURRENT)
End Function

Private Function FileLength( _
    filehandle As HFILE _
) As Long

    FileLength = GetFileSize(filehandle.handle, 0)
End Function

Private Function FileEnd( _
    filehandle As HFILE _
) As Boolean

    FileEnd = FilePosition(filehandle) >= FileLength(filehandle)
End Function
