VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncoderWMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Windows Media Audio 9 Encoder

Implements IEncoder

Private Declare Function CallWindowProc Lib "user32" _
Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private pIProfileManager        As Long
Private IProfileManager         As IWMProfileManager

Private pICodecInfo             As Long
Private ICodecInfo2             As IWMCodecInfo2

Private pIWriter                As Long
Private IWriter                 As IWMWriter

Private pIStreamCfg             As Long
Private IStreamCfg              As IWMStreamConfig

Private pIProfile               As Long
Private IProfile                As IWMProfile

Private lngSelectedCodec        As Long
Private lngSelectedFormat       As Long

Private udtInputWFX             As MMWAVEFORMATEX

Private dblTime                 As Double

Private strTitle                As String
Private strAlbum                As String
Private strArtist               As String

' ***************************************
' * constructor
' ***************************************

Private Sub Class_Initialize()
    Dim guidCodecInfo2  As GUID

    ' create IProfileManager
    If WMCreateProfileManager(pIProfileManager) <> S_OK Then
        Debug.Print "Konnte IWMProfileManager nicht erzeugen"
    Else
        ' VTable of IProfileManager
        DXCopyMemory IProfileManager, ByVal DeRefI4(pIProfileManager), Len(IProfileManager)

        guidCodecInfo2 = Str2GUID(IID_IWMCodecInfo2)

        If CallPtr(IProfileManager.IUnk.QueryInterface, _
                   pIProfileManager, _
                   VarPtr(guidCodecInfo2), _
                   VarPtr(pICodecInfo)) <> S_OK Then

            Debug.Print "Konnte IWMCodecInfo2 nicht erstellen"
        Else
            ' VTable of ICodecInfo
            DXCopyMemory ICodecInfo2, ByVal DeRefI4(pICodecInfo), Len(ICodecInfo2)
        End If
    End If

    lngSelectedCodec = -1
    lngSelectedFormat = -1
End Sub

Private Sub Class_Terminate()
    CallPtr ICodecInfo2.IUnk.Release, pICodecInfo
    CallPtr IProfileManager.IUnk.Release, pIProfileManager
End Sub

' ***************************************
' * Encoder settings
' ***************************************

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
    IEncoder_Extension = "wma"
End Property

Public Function SelectFormat( _
    ByVal codec As Long, _
    ByVal format As Long _
) As Boolean

    If codec < 0 Or codec > CodecsCount Then
        Exit Function
    End If

    If CodecFormatCount(codec) = 0 Then
        Exit Function
    End If

    If format < 0 Or format > CodecFormatCount(codec) Then
        Exit Function
    End If

    lngSelectedCodec = codec
    lngSelectedFormat = format

    SelectFormat = True
End Function

Public Property Get codecname( _
    index As Long _
) As String

    codecname = GetCodecName(index)
End Property

Public Property Get CodecsCount() As Long
    CodecsCount = GetCodecsCount()
End Property

Public Property Get CodecFormatCount( _
    codec As Long _
) As Long

    CodecFormatCount = GetCodecFormatCount(codec)
End Property

Public Property Get CodecFormatName( _
    codec As Long, _
    index As Long _
) As String

    CodecFormatName = TrimEx(GetCodecFormatName(codec, index))
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

    Dim piInpProps  As Long
    Dim IInpProps   As IWMInputMediaProps
    Dim cFormats    As Long
    Dim cbSize      As Long
    Dim btMedia()   As Byte
    Dim guidFormat  As GUID
    Dim udtMedia    As WM_MEDIA_TYPE
    Dim udtWFX      As MMWAVEFORMATEX
    Dim intNumber   As Integer
    Dim i           As Long
    Dim j           As Long

    If lngSelectedCodec = -1 Then _
        Exit Function
    If lngSelectedFormat = -1 Then
        Exit Function
    End If

    If pIWriter <> 0 Then
        IEncoder_Init = STREAM_OK
        Exit Function
    End If

    guidFormat = Str2GUID(WMMEDIATYPE_Audio)

    ' get Stream Config
    If S_OK <> CallPtr(ICodecInfo2.GetCodecFormat, _
                       pICodecInfo, _
                       VarPtr(guidFormat), _
                       lngSelectedCodec, _
                       lngSelectedFormat, _
                       VarPtr(pIStreamCfg)) Then

        IEncoder_Init = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    ' IWMStreamConfig VTable
    DXCopyMemory IStreamCfg, ByVal DeRefI4(pIStreamCfg), Len(IStreamCfg)

    CallPtr IStreamCfg.GetStreamNumber, pIStreamCfg, VarPtr(intNumber)

    If S_OK <> CallPtr(IStreamCfg.SetStreamNumber, pIStreamCfg, 1) Then
        IEncoder_Init = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    ' create a new profile
    If S_OK <> CallPtr(IProfileManager.CreateEmtpyProfile, _
                       pIProfileManager, _
                       WMT_VER_9_0, _
                       VarPtr(pIProfile)) Then

        Exit Function
    End If

    ' IWMProfile VTable
    DXCopyMemory IProfile, ByVal DeRefI4(pIProfile), Len(IProfile)

    ' add stream config to the profile
    If S_OK <> CallPtr(IProfile.AddStream, _
                       pIProfile, _
                       pIStreamCfg) Then

        IEncoder_Init = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    If WMCreateWriter(0, pIWriter) <> S_OK Then
        IEncoder_Init = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    ' IWMWriter VTable
    DXCopyMemory IWriter, ByVal DeRefI4(pIWriter), Len(IWriter)

    If S_OK <> CallPtr(IWriter.SetProfile, _
                       pIWriter, _
                       pIProfile) Then

        IEncoder_Init = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    If S_OK <> CallPtr(IWriter.SetOutputFilename, _
                       pIWriter, _
                       StrPtr(dstFilename)) Then

        IEncoder_Init = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    With udtInputWFX
        .nSamplesPerSec = srcSampleRate
        .nChannels = srcChannels
        .wBitsPerSample = srcBitsPerSample
        .nBlockAlign = .nChannels * (.wBitsPerSample / 8)
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
    End With

    CallPtr IWriter.GetInputFormatCount, _
            pIWriter, _
            0, _
            VarPtr(cFormats)

    For i = 0 To cFormats - 1
        ' IWMInputMediaProps of the current format
        If S_OK <> CallPtr(IWriter.GetInputFormat, _
                           pIWriter, _
                           0, _
                           i, _
                           VarPtr(piInpProps)) Then

            GoTo NextFormat
        End If

        ' IWMInputMediaProps VTable
        DXCopyMemory IInpProps, ByVal DeRefI4(piInpProps), Len(IInpProps)

        ' WM_MEDIA_TYPE of the format
        CallPtr IInpProps.GetMediaType, piInpProps, 0, VarPtr(cbSize)

        ReDim btMedia(cbSize - 1) As Byte

        CallPtr IInpProps.GetMediaType, _
                piInpProps, _
                VarPtr(btMedia(0)), _
                VarPtr(cbSize)

        DXCopyMemory udtMedia, btMedia(0), Len(udtMedia)

        If GUIDEqual(udtMedia.subtype, Str2GUID(WMMEDIASUBTYPE_PCM)) Then
            DXCopyMemory udtWFX, ByVal udtMedia.pbFormat, Len(udtWFX)

            If udtWFX.nSamplesPerSec = srcSampleRate Then
                If udtWFX.nChannels = srcChannels Then
                    If udtWFX.wBitsPerSample = srcBitsPerSample Then

                        ' set the new input format
                        If S_OK <> CallPtr(IInpProps.SetMediaType, _
                                           piInpProps, _
                                           VarPtr(btMedia(0))) Then

                            GoTo NextFormat
                        End If

                        If S_OK <> CallPtr(IWriter.SetInputProps, _
                                           pIWriter, _
                                           0, _
                                           piInpProps) Then

                            GoTo NextFormat
                        End If

                        ' write WM tags
                        WMAAddTag attr_WMAlbumArtist, strArtist
                        WMAAddTag attr_WMAlbumTitle, strAlbum
                        WMAAddTag attr_WMTitle, strTitle

                        ' start the encoder
                        If S_OK = CallPtr(IWriter.BeginWriting, _
                                          pIWriter) Then

                            cbPCMData = 8& * 1024&
                            IEncoder_Init = STREAM_OK
                            Exit Function
                        End If

                    End If
                End If
            End If
        End If

NextFormat:
    Next

    IEncoder_Init = STREAM_NOT_SUPPORTED
End Function

Private Function IEncoder_Encode( _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As STREAM_ERR

    Dim qwTime      As QWORD
    Dim strHexTime  As String
    '
    Dim pIBuffer    As Long
    Dim IBuffer     As INSSBuffer
    Dim pBuffer     As Long
    '
    Dim hr          As HRESULT

    ' new INSSBuffer instance
    hr = CallPtr(IWriter.AllocateSample, _
                 pIWriter, _
                 bytes, _
                 VarPtr(pIBuffer))

    If hr <> S_OK Then
        IEncoder_Encode = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    ' INSSBuffer VTable
    DXCopyMemory IBuffer, ByVal DeRefI4(pIBuffer), Len(IBuffer)

    ' get a pointer to the buffer
    hr = CallPtr(IBuffer.GetBuffer, _
                 pIBuffer, _
                 VarPtr(pBuffer))

    ' fill the buffer
    DXCopyMemory ByVal pBuffer, ByVal ptr, bytes

    ' IWMWriter::WriteSample wants a QWORD,
    ' VB doesn't have one, so simply split the time
    ' in 2 DWORDs (double => hex => LoWord, HiWord)
    strHexTime = FmtStrLen(DecToHex(dblTime), 16)
    qwTime.lo = Val("&H" & Right$(strHexTime, 8) & "&")
    qwTime.hi = Val("&H" & Left$(strHexTime, 8) & "&")

    hr = CallPtr(IWriter.WriteSample, _
                 pIWriter, _
                 0, _
                 qwTime.lo, _
                 qwTime.hi, _
                 0, _
                 pIBuffer)

    If hr = S_OK Then
        IEncoder_Encode = STREAM_OK
    Else
        IEncoder_Encode = STREAM_END_OF_STREAM
    End If

    ' current position in 100 ns
    With udtInputWFX
        dblTime = dblTime + Fix((bytes / .nAvgBytesPerSec) / 10 ^ -7)
    End With

    ' Set IBuffer = Nothing
    CallPtr IBuffer.IUnk.Release, pIBuffer
End Function

Private Function IEncoder_DeInit() As STREAM_ERR
    CallPtr IWriter.EndWriting, pIWriter
    CallPtr IWriter.IUnk.Release, pIWriter
    pIWriter = 0

    CallPtr IProfile.IUnk.Release, pIProfile
    pIProfile = 0

    CallPtr IStreamCfg.IUnk.Release, pIStreamCfg
    pIStreamCfg = 0

    dblTime = 0

    IEncoder_DeInit = STREAM_OK
End Function

' *****************************************
' * helpers
' *****************************************

Private Sub WMAAddTag( _
    ByVal strAttribute As String, _
    ByVal strValue As String _
)

    Dim hr          As HRESULT
    Dim iidHdr      As GUID

    Dim pIHeader    As Long
    Dim IHeader     As IWMHeaderInfo3

    Dim pwIndex     As Integer

    ' get IWMHeaderInfo3 from IWMWriter
    iidHdr = Str2GUID(IID_IWMHeaderInfo3)

    hr = CallPtr(IWriter.IUnk.QueryInterface, _
                 pIWriter, _
                 VarPtr(iidHdr), _
                 VarPtr(pIHeader))

    If hr <> S_OK Then Exit Sub

    ' VTable of IWMHeaderInfo3
    DXCopyMemory IHeader, ByVal DeRefI4(pIHeader), Len(IHeader)

    ' write the attribute
    hr = CallPtr(IHeader.AddAttribute, _
                 pIHeader, _
                 0, _
                 StrPtr(strAttribute), _
                 VarPtr(pwIndex), _
                 WMT_TYPE_STRING, _
                 0, _
                 StrPtr(strValue), _
                 LenB(strValue) + 2)

    ' Set IHeader = Nothing
    CallPtr IHeader.IUnk.Release, pIHeader
End Sub

Private Function GetCodecsCount() As Long
    Dim cCodecs     As Long
    Dim guidFormat  As GUID

    guidFormat = Str2GUID(WMMEDIATYPE_Audio)

    If S_OK = CallPtr(ICodecInfo2.GetCodecInfoCount, _
                      pICodecInfo, _
                      VarPtr(guidFormat), _
                      VarPtr(cCodecs)) Then

        GetCodecsCount = cCodecs
    End If
End Function

Private Function GetCodecName( _
    index As Long _
) As String

    Dim strName     As String
    Dim cChars      As Long
    Dim guidFormat  As GUID

    strName = Space(255)
    cChars = Len(strName)

    guidFormat = Str2GUID(WMMEDIATYPE_Audio)

    CallPtr ICodecInfo2.GetCodecName, _
            pICodecInfo, _
            VarPtr(guidFormat), _
            index, _
            StrPtr(strName), _
            VarPtr(cChars)

    GetCodecName = TrimEx(strName)
End Function

Private Function GetCodecFormatName( _
    codec As Long, _
    index As Long _
) As String

    Dim guidFormat  As GUID
    Dim strFormat   As String
    Dim cChars      As Long

    guidFormat = Str2GUID(WMMEDIATYPE_Audio)

    CallPtr ICodecInfo2.GetCodecFormatDesc, _
            pICodecInfo, _
            VarPtr(guidFormat), _
            codec, _
            index, _
            0, 0, _
            VarPtr(cChars)

    strFormat = Space(cChars)

    CallPtr ICodecInfo2.GetCodecFormatDesc, _
            pICodecInfo, _
            VarPtr(guidFormat), _
            codec, _
            index, _
            0, _
            StrPtr(strFormat), _
            VarPtr(cChars)

    GetCodecFormatName = strFormat
End Function

Private Function GetCodecFormatCount( _
    index As Long _
) As Long

    Dim cFormats    As Long
    Dim guidFormat  As GUID

    guidFormat = Str2GUID(WMMEDIATYPE_Audio)

    If S_OK = CallPtr(ICodecInfo2.GetCodecFormatCount, _
                      pICodecInfo, _
                      VarPtr(guidFormat), _
                      index, _
                      VarPtr(cFormats)) Then

        GetCodecFormatCount = cFormats
    End If
End Function

Private Function CallPtr( _
    ByVal fnc As Long, _
    ParamArray Params() _
) As Long

    Dim btASM(&HEC00& - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    pASM = VarPtr(btASM(0))

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    If UBound(Params) = 0 Then
        If IsArray(Params(0)) Then
            For i = UBound(Params(0)) To 0 Step -1
                AddPush pASM, CLng(Params(0)(i))    ' PUSH dword
            Next
        Else
            For i = UBound(Params) To 0 Step -1
                AddPush pASM, CLng(Params(i))       ' PUSH dword
            Next
        End If
    Else
        For i = UBound(Params) To 0 Step -1
            AddPush pASM, CLng(Params(i))           ' PUSH dword
        Next
    End If

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallPtr = CallWindowProc(VarPtr(btASM(0)), _
                             0, 0, 0, 0)
End Function

Private Sub AddPush( _
    pASM As Long, _
    lng As Long _
)

    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall( _
    pASM As Long, _
    addr As Long _
)

    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong( _
    pASM As Long, _
    lng As Long _
)

    DXCopyMemory ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte( _
    pASM As Long, _
    Bt As Byte _
)

    DXCopyMemory ByVal pASM, Bt, 1
    pASM = pASM + 1
End Sub

Private Function DeRefI4( _
    ByVal ptr As Long _
) As Long

    DXCopyMemory DeRefI4, ByVal ptr, 4
End Function

Private Function Str2GUID( _
    strGUID As String _
) As GUID

    Dim g           As GUID
    Dim strParts()  As String

    strParts = Split(Mid$(strGUID, 2, Len(strGUID) - 2), "-")

    g.Data1 = Val("&H" & strParts(0))
    g.Data2 = Val("&H" & strParts(1))
    g.Data3 = Val("&H" & strParts(2))
    g.Data4(0) = Val("&H" & Mid$(strParts(3), 1, 2))
    g.Data4(1) = Val("&H" & Mid$(strParts(3), 3, 2))
    g.Data4(2) = Val("&H" & Mid$(strParts(4), 1, 2))
    g.Data4(3) = Val("&H" & Mid$(strParts(4), 3, 2))
    g.Data4(4) = Val("&H" & Mid$(strParts(4), 5, 2))
    g.Data4(5) = Val("&H" & Mid$(strParts(4), 7, 2))
    g.Data4(6) = Val("&H" & Mid$(strParts(4), 9, 2))
    g.Data4(7) = Val("&H" & Mid$(strParts(4), 11, 2))

    Str2GUID = g
End Function

Private Function GUIDEqual( _
    guid1 As GUID, _
    guid2 As GUID _
) As Boolean

    If guid1.Data1 = guid2.Data1 Then
        If guid1.Data2 = guid2.Data2 Then
            If guid1.Data3 = guid2.Data3 Then
                If StrConv(guid1.Data4, vbUnicode) = StrConv(guid2.Data4, vbUnicode) Then
                    GUIDEqual = True
                End If
            End If
        End If
    End If
End Function

Private Function TrimEx( _
    ByVal strVal As String _
) As String

    TrimEx = Trim$(Left$(strVal, InStrRev(strVal, Chr$(0)) - 1))
End Function
