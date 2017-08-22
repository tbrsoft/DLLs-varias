VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamCDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Compact Disc Audio Stream (Wrapper for CDManager)

Implements ISoundStream
Implements ISoundInfo

Private Const WAV_RIFF  As Long = 1179011410
Private Const WAV_CDDA  As Long = 1094992963
Private Const WAV_DATA  As Long = 1635017060
Private Const WAV_FMT   As Long = 544501094

Private Type CDAHdr
    RIFF                As Long         ' RIFF Chunk
    hdrlen              As Long         ' Header Length
    CDDA                As Long         ' CDDA Format
    fmt                 As Long         ' fmt Chunk
    something           As Long         ' ?
End Type

Private Type CDAFileHeader
    hdr                 As CDAHdr       ' CDA Header
    ver                 As Integer      ' Version
    NumTrk              As Integer      ' Track Number
    CD_ID               As Long         ' Disc ID
    TrkStartHSG         As Long         ' Track Start  (HSG)
    TrkLenHSG           As Long         ' Track Length (HSG)
    TrkStartRBook(3)    As Byte         ' Track Start  (MSF)
    TrkLenRBook(3)      As Byte         ' Track Length (MSF)
End Type

Private clsStream       As ISoundStream
Private clsStreamInfo   As ISoundInfo
Private clsCDM          As CDManager

Private clsTags         As StreamTags

Private blnReady        As Boolean

' *************************************
' * constructor
' *************************************

Private Sub Class_Initialize()
    Set clsCDM = New CDManager
    Set clsTags = New StreamTags

    Set clsStream = clsCDM
    Set clsStreamInfo = clsCDM

    blnReady = clsCDM.InitManager()
End Sub

Private Sub Class_Terminate()
    clsStream.CloseSource
End Sub

' *************************************
' * drive info
' *************************************

Public Property Get DeviceCount() As Long
    DeviceCount = clsCDM.DeviceCount()
End Property

Public Property Get DeviceName( _
    ByVal index As Long _
) As String

    DeviceName = clsCDM.DeviceName(index)
End Property

Public Property Get DeviceChar( _
    ByVal index As Long _
) As String

    DeviceChar = clsCDM.DeviceChar(index)
End Property

Public Function SelectDevice( _
    ByVal index As Long _
) As Boolean

    SelectDevice = clsCDM.SelectDevice(index)
End Function

Public Property Get toc() As CDTOC
    Set toc = clsCDM.GetTOC()
End Property

Public Property Get DiscPresent() As Boolean
    DiscPresent = clsCDM.IsDiscPresent()
End Property

Public Function OpenCD() As Boolean
    OpenCD = clsCDM.OpenCD()
End Function

Public Function CloseCD() As Boolean
    CloseCD = clsCDM.CloseCD()
End Function

' *************************************
' * Stream properties
' *************************************

Private Property Get ISoundInfo_Force16Bit() As Boolean
    ISoundInfo_Force16Bit = True
End Property

Private Property Get ISoundInfo_Tags() As StreamTags
    Set ISoundInfo_Tags = clsTags
End Property

Private Property Get ISoundStream_EndOfStream() As Boolean
    ISoundStream_EndOfStream = clsStream.EndOfStream
End Property

Private Property Get ISoundStream_Extensions() As String()
    ISoundStream_Extensions = clsStream.Extensions
End Property

Private Function ISoundStream_CloseSource() As STREAM_ERR
    ISoundStream_CloseSource = clsStream.CloseSource()
End Function

Private Property Get ISoundStream_Info() As ISoundInfo
    Set ISoundStream_Info = Me
End Property

Private Property Get ISoundStream_Seekable() As Boolean
    ISoundStream_Seekable = True
End Property

Private Property Get ISoundStream_StreamInfo() As ISoundInfo
    Set ISoundStream_StreamInfo = Me
End Property

Private Property Get ISoundInfo_Bitrate() As Long
    ISoundInfo_Bitrate = clsStreamInfo.Bitrate
End Property

Private Property Get ISoundInfo_BitsPerSample() As Integer
    ISoundInfo_BitsPerSample = clsStreamInfo.bitspersample
End Property

Private Property Get ISoundInfo_Channels() As Integer
    ISoundInfo_Channels = clsStreamInfo.channels
End Property

Private Property Get ISoundInfo_Duration() As Long
    ISoundInfo_Duration = clsStreamInfo.Duration
End Property

Private Property Get ISoundInfo_Position() As Long
    ISoundInfo_Position = clsStreamInfo.position
End Property

Private Property Get ISoundInfo_Samplerate() As Long
    ISoundInfo_Samplerate = clsStreamInfo.samplerate
End Property

Private Property Get ISoundStream_WFXPointer() As Long
    ISoundStream_WFXPointer = clsStream.WFXPointer
End Property

' *************************************
' * main functions
' *************************************

Private Function ISoundStream_Read( _
    ByVal pointer As Long, _
    ByVal cBytes As Long, _
    cBytesRead As Long _
) As STREAM_ERR

    ISoundStream_Read = clsStream.Read(pointer, _
                                       cBytes, _
                                       cBytesRead)
End Function

Private Function ISoundStream_SeekTo( _
    ByVal value As Long, _
    mode As SEEK_MODE _
) As STREAM_ERR

    ISoundStream_SeekTo = clsStream.SeekTo(value, _
                                           mode)
End Function

Private Function ISoundStream_OpenSource( _
    ByVal Source As String _
) As STREAM_ERR

    Dim udtCDAHdr       As CDAFileHeader
    Dim hInp            As hFile
    Dim LBA             As Long
    Dim LBALen          As Long
    Dim i               As Long
    Dim j               As Long
    Dim intFoundTrack   As Integer

    If Not blnReady Then
        ISoundStream_OpenSource = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    If GetDirLevel(Source) > 1 Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    hInp = FileOpen(Source, GENERIC_READ, FILE_SHARE_READ)
    If hInp.handle = INVALID_HANDLE Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    FileRead hInp, VarPtr(udtCDAHdr), Len(udtCDAHdr)

    FileClose hInp

    If udtCDAHdr.hdr.RIFF <> WAV_RIFF Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    If udtCDAHdr.hdr.fmt <> WAV_FMT Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    If udtCDAHdr.hdr.CDDA <> WAV_CDDA Then
        ISoundStream_OpenSource = STREAM_INVALID_SOURCE
        Exit Function
    End If

    Dim udtMSF  As MSF

    udtMSF.m = udtCDAHdr.TrkStartRBook(2)
    udtMSF.s = udtCDAHdr.TrkStartRBook(1)
    udtMSF.F = udtCDAHdr.TrkStartRBook(0)
    LBA = MSF2LBA(udtMSF)

    udtMSF.m = udtCDAHdr.TrkLenRBook(2)
    udtMSF.s = udtCDAHdr.TrkLenRBook(1)
    udtMSF.F = udtCDAHdr.TrkLenRBook(0)
    LBALen = MSF2LBA(udtMSF)

    ' look for the track in each drive,
    ' as we can't rely on the drive letter
    ' (ASPI...)
    For i = 0 To clsCDM.DeviceCount - 1
        If clsCDM.SelectDevice(i) Then
            With clsCDM.GetTOC
                For j = 1 To .TrackCount
                    If .track(j).StartLBA = LBA Then
                        If LBALen = (.track(j + 1).StartLBA - .track(j).StartLBA - 150) Then
                            If clsCDM.SelectTrack(j) = STREAM_OK Then
                                intFoundTrack = 1
                            End If
                        End If
                    End If
                    If intFoundTrack = 1 Then Exit For
                Next
            End With
        End If
        If intFoundTrack = 1 Then Exit For
    Next

    If intFoundTrack < 1 Then
        ISoundStream_OpenSource = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    ISoundStream_OpenSource = STREAM_OK
End Function
