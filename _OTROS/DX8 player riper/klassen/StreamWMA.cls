VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Windows Media Audio/MP3 Decoder
' (needs WMF SDK v9 Runtime or WMP 9/10)

Implements ISoundStream
Implements ISoundInfo

Private Declare Function CallWindowProc Lib "user32" _
Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

' Synchronous WMF Reader
Private IReader                 As IWMSyncReader
Private oReader                 As Long
Private pReaderVTbl             As Long

' Media Properties
Private IProps                  As IWMOutputMediaProps
Private oProps                  As Long
Private pPropsVTbl              As Long

' Media Header
Private IHeader                 As IWMHeaderInfo3
Private oHeader                 As Long
Private pHeaderVTbl             As Long

' waveOut WAVEFORMATEX
Private wfxout                  As MMWAVEFORMATEX

' WMA Attributes
Private curDuration             As Currency
Private curBitrate              As Currency
Private curPosition             As Currency
Private blnSeekable             As Boolean
Private blnEOS                  As Boolean

' PCM Buffer
Private btBuffer()              As Byte
Private lngBufferLen            As Long
Private lngBufferPos            As Long

' WM Core DLL
Private clsWMCore               As clsDLL

Private blnFileOpened           As Boolean

Private clsTags                 As StreamTags

' **********************************
' * constructor
' **********************************

Private Sub Class_Initialize()
    Dim lngRet  As Long

    Set clsWMCore = New clsDLL
    Set clsTags = New StreamTags

    If Not clsWMCore.LoadDLL("wmvcore.dll") Then
        Exit Sub
    End If

    ' create IWMSyncReader instance
    lngRet = clsWMCore.CallFunc("WMCreateSyncReader", 0, WMT_RIGHT_PLAYBACK, VarPtr(oReader))
    If lngRet <> 0 Then
        oReader = 0
        Exit Sub
    End If

    ' IWMSyncReader`s VTable
    DXCopyMemory pReaderVTbl, ByVal oReader, 4
    DXCopyMemory IReader, ByVal pReaderVTbl, Len(IReader)
End Sub

Private Sub Class_Terminate()
    Dim lngRet  As Long

    ISoundStream_CloseSource

    ' Set IWMSyncReader = Nothing
    If oReader <> 0 Then
        CallPointer IReader.IUnk.Release, oReader
    End If

    clsWMCore.UnloadDLL
End Sub

' **********************************
' * Stream properties
' **********************************

Private Property Get ISoundInfo_Bitrate() As Long
    ISoundInfo_Bitrate = CLng(curBitrate * 10000)
End Property

Private Property Get ISoundInfo_BitsPerSample() As Integer
    ISoundInfo_BitsPerSample = 16
End Property

Private Property Get ISoundInfo_Channels() As Integer
    ISoundInfo_Channels = wfxout.nChannels
End Property

Private Property Get ISoundInfo_Duration() As Long
    ISoundInfo_Duration = curDuration
End Property

Private Property Get ISoundInfo_Force16Bit() As Boolean
    ISoundInfo_Force16Bit = True
End Property

Private Property Get ISoundInfo_Position() As Long
    ISoundInfo_Position = curPosition
End Property

Private Property Get ISoundInfo_Samplerate() As Long
    ISoundInfo_Samplerate = wfxout.nSamplesPerSec
End Property

Private Property Get ISoundInfo_Tags() As StreamTags
    Set ISoundInfo_Tags = clsTags
End Property

Private Property Get ISoundStream_EndOfStream() As Boolean
    ISoundStream_EndOfStream = blnEOS
End Property

Private Property Get ISoundStream_Extensions() As String()
    Dim strExt(1)   As String

    strExt(0) = "wma"
    strExt(1) = "mp3"

    ISoundStream_Extensions = strExt
End Property

Private Property Get ISoundStream_Info() As ISoundInfo
    Set ISoundStream_Info = Me
End Property

Private Property Get ISoundStream_Seekable() As Boolean
    ISoundStream_Seekable = blnSeekable
End Property

Private Property Get ISoundStream_WFXPointer() As Long
    ISoundStream_WFXPointer = VarPtr(wfxout)
End Property

' **********************************
' * main functions
' **********************************

Private Function ISoundStream_CloseSource() As STREAM_ERR
    Dim lngRet  As Long

    If oReader = 0 Then
        ISoundStream_CloseSource = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    If Not blnFileOpened Then
        ISoundStream_CloseSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    lngRet = CallPointer(IReader.Close, oReader)
    blnFileOpened = False

    ISoundStream_CloseSource = STREAM_OK
End Function

Private Function ISoundStream_OpenSource( _
    ByVal Source As String _
) As STREAM_ERR

    Dim pcFormats   As Long, pcOutputs  As Long
    Dim i           As Long, j          As Long
    '+
    '+
    Dim MediaType   As WM_MEDIA_TYPE
    Dim pcbType     As Long, btMT()     As Byte
    '+
    '+
    Dim cAttrs      As Long, attIndex   As Integer
    Dim pwszName    As String, pbValue() As Byte
    Dim strValue    As String
    Dim cchName     As Integer, cbValue  As Long
    Dim attType     As WMT_ATTR_DATATYPE
    '+
    '+
    Dim iid_hdr     As GUID
    '+
    '+
    Dim lngRet      As Long

    ' we got an IWMSyncReader instance?
    If oReader = 0 Then
        ISoundStream_OpenSource = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    If Not IsValidFile(Source) Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    lngRet = CallPointer(IReader.Open, _
                         oReader, _
                         StrPtr(Source))

    If lngRet <> 0 Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    lngRet = CallPointer(IReader.GetOutputCount, _
                         oReader, _
                         VarPtr(pcOutputs))

    If lngRet <> 0 Then
        lngRet = CallPointer(IReader.Close, oReader)
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    For i = 0 To pcOutputs - 1

        lngRet = CallPointer(IReader.GetOutputFormatCount, _
                             oReader, _
                             i, _
                             VarPtr(pcFormats))

        If lngRet <> 0 Then GoTo NextOutput

        For j = 0 To pcFormats

            ' create IWMOutputMediaProps
            lngRet = CallPointer(IReader.GetOutputFormat, _
                                 oReader, _
                                 j, i, _
                                 VarPtr(oProps))

            If lngRet <> 0 Then GoTo NextFormat

            ' IWMOutputMediaProps`s VTable
            DXCopyMemory pPropsVTbl, ByVal oProps, 4
            DXCopyMemory IProps, ByVal pPropsVTbl, Len(IProps)

            ' it is not allowed to directly get
            ' WM_MEDIA_TYPE.
            ' You first have to get the size of WM_MEDIA_TYPE,
            ' allocate some mem for it and then
            ' call GetMediaType again.

            lngRet = CallPointer(IProps.GetMediaType, _
                                 oProps, _
                                 0, _
                                 VarPtr(pcbType))

            If lngRet <> 0 Then
                lngRet = CallPointer(IProps.IUnk.Release, oProps)
                oProps = 0
                GoTo NextFormat
            Else
                ReDim btMT(pcbType - 1) As Byte
            End If

            lngRet = CallPointer(IProps.GetMediaType, _
                                 oProps, _
                                 VarPtr(btMT(0)), _
                                 VarPtr(pcbType))

            DXCopyMemory MediaType, btMT(0), Len(MediaType)

            ' Audio?
            If GUIDEqual(MediaType.majortype, Str2GUID(WMMEDIATYPE_Audio)) Then
                ' PCM?
                If GUIDEqual(MediaType.subtype, Str2GUID(WMMEDIASUBTYPE_PCM)) Then
                    ' 16 Bit?
                    DXCopyMemory wfxout, ByVal MediaType.pbFormat, Len(wfxout)
                    If wfxout.wBitsPerSample = 16 Then
                        lngRet = CallPointer(IProps.SetMediaType, oProps, VarPtr(btMT(0)))
                        GoTo ExitLoops
                    End If
                End If
            End If

            ' Set IWMOutputMediaProps = Nothing
            lngRet = CallPointer(IProps.IUnk.Release, oProps)
            oProps = 0

NextFormat:
        Next j

NextOutput:
    Next i
ExitLoops:

    If i = pcOutputs Then
        ISoundStream_OpenSource = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    ' set the new output format
    lngRet = CallPointer(IReader.SetOutputProps, oReader, i, oProps)
    lngRet = CallPointer(IProps.IUnk.Release, oProps)
    oProps = 0

    ' region to decode (0 to EOF)
    lngRet = CallPointer(IReader.SetRangeByFrame, _
                         oReader, _
                         1, 0, 0, _
                         0, 0)

    If lngRet <> 0 Then
        lngRet = CallPointer(IReader.Close, oReader)
        ISoundStream_OpenSource = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    ' get IWMHeaderInfo3 from IWMSyncReader
    iid_hdr = Str2GUID(IID_IWMHeaderInfo3)
    lngRet = CallPointer(IReader.IUnk.QueryInterface, _
                         oReader, _
                         VarPtr(iid_hdr), _
                         VarPtr(oHeader))

    If lngRet <> 0 Then Exit Function

    ' IWMHeaderInfo3`s VTable
    DXCopyMemory pHeaderVTbl, ByVal oHeader, 4
    DXCopyMemory IHeader, ByVal pHeaderVTbl, Len(IHeader)

    lngRet = CallPointer(IHeader.GetAttributeCountEx, _
                         oHeader, _
                         &HFFFF&, _
                         VarPtr(cAttrs))

    If lngRet <> 0 Then
        lngRet = CallPointer(IHeader.IUnk.Release, oHeader)
        ISoundStream_OpenSource = STREAM_OK
        oHeader = 0
        Exit Function
    End If

    GetID3Tags Source

    For attIndex = 0 To cAttrs - 1

        lngRet = CallPointer(IHeader.GetAttributeByIndexEx, _
                             oHeader, _
                             &HFFFF&, _
                             attIndex, _
                             0, _
                             VarPtr(cchName), _
                             0, 0, 0, _
                             VarPtr(cbValue))

        If lngRet <> 0 Then GoTo NextAttribute

        ' buffer for attribute's name and value
        pwszName = Space(cchName)
        ReDim pbValue(cbValue - 1) As Byte

        lngRet = CallPointer(IHeader.GetAttributeByIndexEx, _
                             oHeader, _
                             &HFFFF&, _
                             attIndex, _
                             StrPtr(pwszName), _
                             VarPtr(cchName), _
                             VarPtr(attType), _
                             0, _
                             VarPtr(pbValue(0)), _
                             VarPtr(cbValue))

        If lngRet <> 0 Then GoTo NextAttribute

        pwszName = RemNullChars(pwszName)

        Select Case attType
            Case WMT_TYPE_WORD                  ' integer
                Dim intVal  As Integer
                DXCopyMemory intVal, pbValue(0), 2
                strValue = intVal
            Case WMT_TYPE_DWORD                 ' long
                Dim lngVal  As Long
                DXCopyMemory lngVal, pbValue(0), 4
                strValue = lngVal
            Case WMT_TYPE_QWORD                 ' currency?
                Dim curVal  As Currency
                DXCopyMemory curVal, pbValue(0), 8
                strValue = curVal
            Case WMT_TYPE_STRING                ' Unicode String
                strValue = RemNullChars(pbValue)
            Case WMT_TYPE_BINARY                ' binary...
                ' nicht readable...
            Case WMT_TYPE_BOOL                  ' boolean
                DXCopyMemory lngVal, pbValue(0), cbValue
                strValue = CBool(lngVal)
            Case WMT_TYPE_GUID                  ' global unique identifier
                Dim g   As GUID
                DXCopyMemory g, pbValue(0), Len(g)
                strValue = GUID2Str(g)
        End Select

        clsTags.AddTag pwszName, strValue

        Select Case pwszName

            Case attr_WMDuration
                CpyAttribute VarPtr(curDuration), _
                             VarPtr(pbValue(0)), _
                             attType

            Case attr_WMBitrate
                CpyAttribute VarPtr(curBitrate), _
                             VarPtr(pbValue(0)), _
                             attType

            Case attr_WMSeekable
                CpyAttribute VarPtr(blnSeekable), _
                             VarPtr(pbValue(0)), _
                             attType

        End Select

NextAttribute:
    Next

    ' Set IWMHeaderInfo3 = Nothing
    lngRet = CallPointer(IHeader.IUnk.Release, oHeader)
    oHeader = 0

    blnFileOpened = True

    ISoundStream_OpenSource = STREAM_OK
End Function

Private Function ISoundStream_Read( _
    ByVal pBuffer As Long, _
    ByVal cBytes As Long, _
    cBytesRead As Long _
) As STREAM_ERR

    Do While cBytesRead < cBytes

        If lngBufferLen - lngBufferPos < cBytes - cBytesRead Then

            If 0 < (lngBufferLen - lngBufferPos) Then
                If 0 = IsBadWritePtr(ByVal pBuffer + cBytesRead, lngBufferLen - lngBufferPos) Then
                    If 0 = IsBadReadPtr(btBuffer(lngBufferPos), lngBufferLen - lngBufferPos) Then
                        DXCopyMemory ByVal pBuffer + cBytesRead, btBuffer(lngBufferPos), lngBufferLen - lngBufferPos
                    End If
                End If

                cBytesRead = cBytesRead + (lngBufferLen - lngBufferPos)
            End If

            If Not FillBuffer Then
                ISoundStream_Read = STREAM_END_OF_STREAM
                Exit Function
            End If

        Else

            If 0 = IsBadWritePtr(ByVal pBuffer + cBytesRead, cBytes - cBytesRead) Then
                If 0 = IsBadReadPtr(btBuffer(lngBufferPos), cBytes - cBytesRead) Then
                    DXCopyMemory ByVal pBuffer + cBytesRead, btBuffer(lngBufferPos), cBytes - cBytesRead
                End If
            End If

            lngBufferPos = lngBufferPos + (cBytes - cBytesRead)
            cBytesRead = cBytesRead + (cBytes - cBytesRead)

        End If

    Loop

    ISoundStream_Read = STREAM_OK
End Function

Private Function ISoundStream_SeekTo( _
    ByVal value As Long, _
    mode As SEEK_MODE _
) As STREAM_ERR

    Dim curTime     As QWORD, curTime0  As QWORD
    '
    Dim lngRet      As Long
    '
    Dim dblTime     As Double, strHex   As String


    ' IWMSyncReader::SetRange wants QWORDS.
    ' VB doesn't support them,
    ' so we calculate the value in a double,
    ' convert it to hex, split it in 2 DWORDS,
    ' and put them on the stack.

    Select Case mode

        Case SEEK_PERCENT
            If value < 0 Or value > 100 Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            dblTime = Fix(curDuration / CDbl(100) * CDbl(value) * 10000#)

        Case SEEK_SECONDS
            If value < 0 Or value > (ISoundInfo_Duration / 1000) Then
                ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
                Exit Function
            End If

            dblTime = Fix(curDuration / CDbl(100) * CDbl(value / (ISoundInfo_Duration / 1000) * 100) * 10000#)

    End Select

    strHex = FmtStrLen(DecToHex(dblTime), 16)

    ' lo DWORD
    curTime.lo = Val("&H" & Right$(strHex, 8) & "&")
    ' hi DWORD
    curTime.hi = Val("&H" & Left$(strHex, 8) & "&")

    lngRet = CallPointer(IReader.SetRange, oReader, _
                         curTime.lo, curTime.hi, _
                         curTime0.lo, curTime0.hi)

    If lngRet = 0 Then
        FillBuffer
        ISoundStream_SeekTo = STREAM_OK
    Else
        ISoundStream_SeekTo = STREAM_OUT_OF_RANGE
    End If
End Function

' **********************************
' * helpers
' **********************************

Private Function FillBuffer() As Boolean
    Dim lngRet          As Long
    ' +
    ' +
    Dim ISample         As INSSBuffer
    Dim oSample         As Long
    Dim pSampleVTbl     As Long
    ' +
    ' +
    Dim curSampleDur    As Currency
    Dim pdwFlags        As Long
    Dim pdwOutputNum    As Long
    ' +
    ' +
    Dim ppdwBuffer      As Long
    Dim pdwLength       As Long


    If oReader = 0 Then Exit Function

    lngRet = CallPointer(IReader.GetNextSample, _
                         oReader, _
                         0, _
                         VarPtr(oSample), _
                         VarPtr(curPosition), _
                         VarPtr(curSampleDur), _
                         VarPtr(pdwFlags), _
                         VarPtr(pdwOutputNum), _
                         0)

    If lngRet <> 0 Then
        blnEOS = lngRet = NS_E_NO_MORE_SAMPLES
        Exit Function
    End If

    ' INSSBuffer's VTable
    DXCopyMemory pSampleVTbl, ByVal oSample, 4
    DXCopyMemory ISample, ByVal pSampleVTbl, Len(ISample)

    ' get a pointer to the PCM data and its length
    lngRet = CallPointer(ISample.GetBufferAndLength, _
                         oSample, _
                         VarPtr(ppdwBuffer), _
                         VarPtr(pdwLength))

    If lngRet <> 0 Then
        lngRet = CallPointer(ISample.IUnk.Release, oSample)
        Exit Function
    End If

    ' copy PCM data to our buffer
    ReDim btBuffer(pdwLength - 1) As Byte
    DXCopyMemory btBuffer(0), ByVal ppdwBuffer, pdwLength

    ' Set INSSBuffer = Nothing
    lngRet = CallPointer(ISample.IUnk.Release, oSample)

    ' reset Buffer Descriptors
    lngBufferPos = 0
    lngBufferLen = pdwLength

    FillBuffer = True
End Function

Private Function CallPointer( _
    ByVal fnc As Long, _
    ParamArray Params() _
) As Long

    Dim btASM(&HEC00& - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    If fnc = 0 Then Exit Function

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

    CallPointer = CallWindowProc(VarPtr(btASM(0)), _
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

Private Sub CpyAttribute( _
    pDst As Long, _
    pSrc As Long, _
    dt As WMT_ATTR_DATATYPE _
)

    Dim nSize   As Long

    Select Case dt
        Case WMT_TYPE_BOOL: nSize = 2
        Case WMT_TYPE_DWORD: nSize = 4
        Case WMT_TYPE_GUID: nSize = 16
        Case WMT_TYPE_QWORD: nSize = 8
        Case WMT_TYPE_WORD: nSize = 2
    End Select

    DXCopyMemory ByVal pDst, ByVal pSrc, nSize
End Sub

Private Function GUID2Str( _
    g As GUID _
) As String

    Dim nTemp   As String

    nTemp = "{"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data1), 8) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data2), 4) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data3), 4) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(0)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(1)), 2) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(2)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(3)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(4)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(5)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(6)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(7)), 2) & "}"

    GUID2Str = nTemp
End Function

Private Function GUIDEqual( _
    g1 As GUID, _
    g2 As GUID _
) As Boolean

    GUIDEqual = GUID2Str(g1) = GUID2Str(g2)
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

Private Sub GetID3Tags( _
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
