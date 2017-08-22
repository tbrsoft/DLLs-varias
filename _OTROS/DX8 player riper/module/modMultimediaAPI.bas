Attribute VB_Name = "modMultimediaAPI"
Option Explicit

' collected Win MM API

Public Declare Function mmioClose Lib "winmm" ( _
    ByVal hmmio As Long, _
    ByVal uFlags As Long _
) As Long

Public Declare Function mmioDescend Lib "winmm" ( _
    ByVal hmmio As Long, _
    lpck As MMCKINFO, _
    lpckParent As MMCKINFO, _
    ByVal uFlags As Long _
) As Long

Public Declare Function mmioDescendParent Lib "winmm" _
Alias "mmioDescend" ( _
    ByVal hmmio As Long, _
    lpck As MMCKINFO, _
    ByVal x As Long, _
    ByVal uFlags As Long _
) As Long

Public Declare Function mmioOpen Lib "winmm" _
Alias "mmioOpenA" ( _
    ByVal szFileName As String, _
    lpmmioinfo As MMIOINFO, _
    ByVal dwOpenFlags As Long _
) As Long

Public Declare Function mmioSeek Lib "winmm" ( _
    ByVal hmmio As Long, _
    ByVal lOffset As Long, _
    ByVal iOrigin As Long _
) As Long

Public Declare Function mmioStringToFOURCC Lib "winmm" _
Alias "mmioStringToFOURCCA" ( _
    ByVal sz As String, _
    ByVal uFlags As Long _
) As Long

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Declare Function acmStreamPrepareHeader Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwPrepare As Long _
) As Long

Public Declare Function acmStreamUnprepareHeader Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwUnprepare As Long _
) As Long

Public Declare Function acmStreamOpen Lib "msacm32" ( _
    phas As Long, _
    ByVal had As Long, _
    pwfxSrc As Any, _
    pwfxDst As Any, _
    ByVal pwfltr As Long, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal fdwOpen As Long _
) As Long

Public Declare Function acmStreamSize Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal cbInput As Long, _
    pdwOutputBytes As Long, _
    ByVal fdwSize As Long _
) As Long

Public Declare Function acmStreamConvert Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwConvert As Long _
) As Long

Public Declare Function acmStreamReset Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal fdwReset As Long _
) As Long

Public Declare Function acmStreamClose Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal fdwClose As Long _
) As Long

Public Declare Function acmFormatChoose Lib "msacm32" _
Alias "acmFormatChooseA" ( _
    pfmtc As ACMFORMATCHOOSEA _
) As Long

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Declare Sub ZeroMemory Lib "kernel32" _
Alias "RtlZeroMemory" ( _
    pDst As Any, _
    ByVal dwLen As Long _
)

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Declare Function IsBadReadPtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Public Declare Function IsBadWritePtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Public Const MMIO_READ          As Long = &H0
Public Const MMIO_FINDCHUNK     As Long = &H10
Public Const MMIO_FINDRIFF      As Long = &H20

Public Const SEEK_CUR           As Long = 1

Public Const ACMFORMATDETAILS_FORMAT_CHARS  As Long = 128
Public Const ACMFORMATTAGDETAILS_FORMATTAG_CHARS As Long = 48

Public Const MM_ACM_FORMATCHOOSE            As Long = &H8000

Public Const FORMATCHOOSE_MESSAGE           As Long = 0
Public Const FORMATCHOOSE_FORMATTAG_VERIFY  As Long = (FORMATCHOOSE_MESSAGE + 0)
Public Const FORMATCHOOSE_FORMAT_VERIFY     As Long = (FORMATCHOOSE_MESSAGE + 1)
Public Const FORMATCHOOSE_CUSTOM_VERIFY     As Long = (FORMATCHOOSE_MESSAGE + 2)

Public Const ACMFORMATCHOOSE_STYLEF_SHOWHELP             As Long = &H4
Public Const ACMFORMATCHOOSE_STYLEF_ENABLEHOOK           As Long = &H8
Public Const ACMFORMATCHOOSE_STYLEF_ENABLETEMPLATE       As Long = &H10
Public Const ACMFORMATCHOOSE_STYLEF_ENABLETEMPLATEHANDLE As Long = &H20
Public Const ACMFORMATCHOOSE_STYLEF_INITTOWFXSTRUCT      As Long = &H40
Public Const ACMFORMATCHOOSE_STYLEF_CONTEXTHELP          As Long = &H80

Public Const ACM_FORMATENUMF_CONVERT    As Long = &H100000

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Type MMWAVEFORMATEX
    wFormatTag                  As Integer
    nChannels                   As Integer
    nSamplesPerSec              As Long
    nAvgBytesPerSec             As Long
    nBlockAlign                 As Integer
    wBitsPerSample              As Integer
    cbSize                      As Integer
End Type

Public Type ACMSTREAMHEADER
    cbStruct                    As Long
    fdwStatus                   As Long
    dwUser                      As Long
    pbSrc                       As Long
    cbSrcLength                 As Long
    cbSrcLengthUsed             As Long
    dwSrcUser                   As Long
    pbDst                       As Long
    cbDstLength                 As Long
    cbDstLengthUsed             As Long
    dwDstUser                   As Long
    dwReservedDriver(9)         As Long
End Type

Public Type ACMFORMATCHOOSEA
    cbStruct                    As Long
    fdwStyle                    As Long
    hwndOwner                   As Long
    pwfx                        As Long
    cbwfx                       As Long
    pszTitle                    As Long
    szFormatTag                 As String * ACMFORMATTAGDETAILS_FORMATTAG_CHARS
    szFormat                    As String * ACMFORMATDETAILS_FORMAT_CHARS
    pszName                     As Long
    cchName                     As Long
    fdwEnum                     As Long
    pwfxEnum                    As Long
    hInstance                   As Long
    pszTemplateName             As Long
    lCustData                   As Long
    pfnHook                     As Long
    btSpace(1023)               As Byte ' zur Sicherheit...
End Type

Public Type MMIOINFO
   dwFlags                      As Long
   fccIOProc                    As Long
   pIOProc                      As Long
   wErrorRet                    As Long
   htask                        As Long
   cchBuffer                    As Long
   pchBuffer                    As String
   pchNext                      As String
   pchEndRead                   As String
   pchEndWrite                  As String
   lBufOffset                   As Long
   lDiskOffset                  As Long
   adwInfo(4)                   As Long
   dwReserved1                  As Long
   dwReserved2                  As Long
   hmmio                        As Long
End Type

Public Type WAVE_FORMAT
    wFormatTag                  As Integer
    wChannels                   As Integer
    dwSampleRate                As Long
    dwBytesPerSec               As Long
    wBlockAlign                 As Integer
    wBitsPerSample              As Integer
End Type

Public Type MMCKINFO
   ckid                         As Long
   ckSize                       As Long
   fccType                      As Long
   dwDataOffset                 As Long
   dwFlags                      As Long
End Type

Public Type CHUNKINFO
    Start                       As Long
    Length                      As Long
End Type

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Enum HACMSTREAM
    INVALID_STREAM_HANDLE = 0
End Enum

Public Enum ACM_STREAMSIZEF
    ACM_STREAMSIZEF_DESTINATION = &H1
    ACM_STREAMSIZEF_SOURCE = &H0
    ACM_STREAMSIZEF_QUERYMASK = &HF
End Enum

Public Enum ACM_STREAMCONVERTF
    ACM_STREAMCONVERTF_BLOCKALIGN = &H4
    ACM_STREAMCONVERTF_START = &H10
    ACM_STREAMCONVERTF_END = &H20
End Enum

' create a PCM WAVEFORMATEX structure
Public Function CreateWFX( _
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

' find a chunk in a WAV container
Public Function GetWavChunkPos( _
    ByVal strFile As String, _
    ByVal strChunk As String _
) As CHUNKINFO

    Dim hMmioIn             As Long
    Dim lngRet              As Long
    Dim mmckinfoParentIn    As MMCKINFO
    Dim mmckinfoSubchunkIn  As MMCKINFO
    Dim mmioinf             As MMIOINFO

    ' open WAV for read access
    hMmioIn = mmioOpen(strFile, mmioinf, MMIO_READ)
    If hMmioIn = 0 Then
        Exit Function
    End If

    ' check for a valid WAV
    mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
    lngRet = mmioDescendParent(hMmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
    If Not (lngRet = 0) Then
        mmioClose hMmioIn, 0
        Exit Function
    End If

    ' search for the chunk
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC(strChunk, 0)
    lngRet = mmioDescend(hMmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
    If Not (lngRet = 0) Then
        mmioClose hMmioIn, 0
        Exit Function
    End If

    ' return the position and length of the chunk
    GetWavChunkPos.Start = mmioSeek(hMmioIn, 0, SEEK_CUR)
    GetWavChunkPos.Length = mmckinfoSubchunkIn.ckSize

    mmioClose hMmioIn, 0
End Function
