Attribute VB_Name = "modWindowsMedia"
Option Explicit

' WMF SDK Header translation (not complete)

Public Declare Function WMCreateProfileManager Lib "wmvcore" ( _
    ppProfile As Any _
) As HRESULT

Public Declare Function WMCreateWriter Lib "wmvcore" ( _
    ByVal pUnkReserved As Long, _
    ppWriter As Any _
) As HRESULT

Public Enum HRESULT
    S_OK = 0
End Enum

Public Enum WMT_VERSION
    WMT_VER_4_0 = &H40000
    WMT_VER_7_0 = &H70000
    WMT_VER_8_0 = &H80000
    WMT_VER_9_0 = &H90000
End Enum

Public Enum WMT_ATTR_DATATYPE
    WMT_TYPE_DWORD = 0
    WMT_TYPE_STRING = 1
    WMT_TYPE_BINARY = 2
    WMT_TYPE_BOOL = 3
    WMT_TYPE_QWORD = 4
    WMT_TYPE_WORD = 5
    WMT_TYPE_GUID = 6
End Enum

Public Enum WMT_ERRORS
    NS_E_NO_MORE_SAMPLES = &HC00D0BCF
End Enum

Public Enum WMT_RIGHTS
    WMT_RIGHT_PLAYBACK = &H1
    WMT_RIGHT_COPY_TO_NON_SDMI_DEVICE = &H2
    WMT_RIGHT_COPY_TO_CD = &H8
    WMT_RIGHT_COPY_TO_SDMI_DEVICE = &H10
    WMT_RIGHT_ONE_TIME = &H20
    WMT_RIGHT_SAVE_STREAM_PROTECTED = &H40
    WMT_RIGHT_COPY = &H80
    WMT_RIGHT_COLLABORATIVE_PLAY = &H100
    WMT_RIGHT_SDMI_TRIGGER = &H10000
    WMT_RIGHT_SDMI_NOMORECOPIES = &H20000
End Enum

Public Type QWORD
    lo                          As Long
    hi                          As Long
End Type

Public Type GUID
    Data1                       As Long
    Data2                       As Integer
    Data3                       As Integer
    Data4(7)                    As Byte
End Type

Public Type IUnknown
    QueryInterface              As Long
    AddRef                      As Long
    Release                     As Long
End Type

Public Type IWMWriter
    IUnk                        As IUnknown
    SetProfileByID              As Long
    SetProfile                  As Long
    SetOutputFilename           As Long
    GetInputCount               As Long
    GetInputProps               As Long
    SetInputProps               As Long
    GetInputFormatCount         As Long
    GetInputFormat              As Long
    BeginWriting                As Long
    EndWriting                  As Long
    AllocateSample              As Long
    WriteSample                 As Long
    Flush                       As Long
End Type

Public Type INSSBuffer
    IUnk                        As IUnknown
    GetLength                   As Long
    SetLength                   As Long
    GetMaxLength                As Long
    GetBuffer                   As Long
    GetBufferAndLength          As Long
End Type

Public Type IWMHeaderInfo
    IUnk                        As IUnknown
    GetAttributeCount           As Long
    GetAttributeByIndex         As Long
    GetAttributeByName          As Long
    SetAttribute                As Long
    GetMarkerCount              As Long
    GetMarker                   As Long
    AddMarker                   As Long
    RemoveMarker                As Long
    GetScriptCount              As Long
    GetScript                   As Long
    AddScript                   As Long
    RemoveScript                As Long
End Type

Public Type IWMHeaderInfo2 ' : IWMHeaderInfo
    IUnk                        As IUnknown
    GetAttributeCount           As Long
    GetAttributeByIndex         As Long
    GetAttributeByName          As Long
    SetAttribute                As Long
    GetMarkerCount              As Long
    GetMarker                   As Long
    AddMarker                   As Long
    RemoveMarker                As Long
    GetScriptCount              As Long
    GetScript                   As Long
    AddScript                   As Long
    RemoveScript                As Long
    GetCodecInfoCount           As Long
    GetCodecInfo                As Long
End Type

Public Type IWMHeaderInfo3 ' : IWMHeaderInfo2
    IUnk                        As IUnknown
    GetAttributeCount           As Long
    GetAttributeByIndex         As Long
    GetAttributeByName          As Long
    SetAttribute                As Long
    GetMarkerCount              As Long
    GetMarker                   As Long
    AddMarker                   As Long
    RemoveMarker                As Long
    GetScriptCount              As Long
    GetScript                   As Long
    AddScript                   As Long
    RemoveScript                As Long
    GetCodecInfoCount           As Long
    GetCodecInfo                As Long
    GetAttributeCountEx         As Long
    GetAttributeIndices         As Long
    GetAttributeByIndexEx       As Long
    ModifyAttribute             As Long
    AddAttribute                As Long
    DeleteAttribute             As Long
    AddCodecInfo                As Long
End Type

Public Type IWMMediaProps
    IUnk                        As IUnknown
    GetType                     As Long
    GetMediaType                As Long
    SetMediaType                As Long
End Type

Public Type IWMInputMediaProps ' : IWMMediaProps
    IUnk                        As IUnknown
    GetType                     As Long
    GetMediaType                As Long
    SetMediaType                As Long
    GetConnectionName           As Long
    GetGroupName                As Long
End Type

Public Type IWMProfileManager
    IUnk                        As IUnknown
    CreateEmtpyProfile          As Long
    LoadProfileByID             As Long
    LoadProfileByData           As Long
    SaveProfile                 As Long
    GetSystemProfileCount       As Long
    LoadSystemProfile           As Long
End Type

Public Type IWMProfile
    IUnk                        As IUnknown
    GetVersion                  As Long
    GetName                     As Long
    SetName                     As Long
    GetDescription              As Long
    SetDescription              As Long
    GetStreamCount              As Long
    GetStream                   As Long
    GetStreamByNumber           As Long
    RemoveStream                As Long
    RemoveStreamByNumber        As Long
    AddStream                   As Long
    ReconfigStream              As Long
    CreateNewStream             As Long
    GetMutualExclusionCount     As Long
    AddMutualExclusion          As Long
    CreateNewMutualExclusion    As Long
End Type

Public Type IWMProfile2 ' : IWMProfile
    IUnk                        As IUnknown
    GetVersion                  As Long
    GetName                     As Long
    SetName                     As Long
    GetDescription              As Long
    SetDescription              As Long
    GetStreamCount              As Long
    GetStream                   As Long
    GetStreamByNumber           As Long
    RemoveStream                As Long
    RemoveStreamByNumber        As Long
    AddStream                   As Long
    ReconfigStream              As Long
    CreateNewStream             As Long
    GetMutualExclusionCount     As Long
    AddMutualExclusion          As Long
    CreateNewMutualExclusion    As Long
    GetProfileID                As Long
End Type

Public Type IWMProfile3
    IUnk                        As IUnknown
    GetVersion                  As Long
    GetName                     As Long
    SetName                     As Long
    GetDescription              As Long
    SetDescription              As Long
    GetStreamCount              As Long
    GetStream                   As Long
    GetStreamByNumber           As Long
    RemoveStream                As Long
    RemoveStreamByNumber        As Long
    AddStream                   As Long
    ReconfigStream              As Long
    CreateNewStream             As Long
    GetMutualExclusionCount     As Long
    AddMutualExclusion          As Long
    CreateNewMutualExclusion    As Long
    GetProfileID                As Long
    GetStorageFormat            As Long
    SetStorageFormat            As Long
    GetBandwidthSharingCount    As Long
    GetBandwidthSharing         As Long
    RemoveBandwidthSharing      As Long
    AddBandwidthSharing         As Long
    CreateNewBandwidthSharing   As Long
    GetStreamPrioritization     As Long
    SetStreamPrioritization     As Long
    RemoveStreamPrioritization  As Long
    CreateNewStreamPrioritization As Long
    GetExpectedPacketCount      As Long
End Type

Public Type IWMCodecInfo
    IUnk                        As IUnknown
    GetCodecInfoCount           As Long
    GetCodecFormatCount         As Long
    GetCodecFormat              As Long
End Type

Public Type IWMCodecInfo2
    IUnk                        As IUnknown
    GetCodecInfoCount           As Long
    GetCodecFormatCount         As Long
    GetCodecFormat              As Long
    GetCodecName                As Long
    GetCodecFormatDesc          As Long
End Type

Public Type IWMStreamConfig
    IUnk                        As IUnknown
    GetStreamType               As Long
    GetStreamNumber             As Long
    SetStreamNumber             As Long
    GetStreamName               As Long
    SetStreamName               As Long
    GetConnectionName           As Long
    SetConnectionName           As Long
    GetBitrate                  As Long
    SetBitrate                  As Long
    GetBufferWindow             As Long
    SetBufferWindow             As Long
End Type

Public Type IWMOutputMediaProps
    IUnk                        As IUnknown
    GetType                     As Long
    GetMediaType                As Long
    SetMediaType                As Long
    GetStreamGroupName          As Long
    GetConnectionName           As Long
End Type

Public Type IWMSyncReader
    IUnk                        As IUnknown
    Open                        As Long
    Close                       As Long
    SetRange                    As Long
    SetRangeByFrame             As Long
    GetNextSample               As Long
    SetStreamsSelected          As Long
    GetStreamSelected           As Long
    SetReadStreamSamples        As Long
    GetReadStreamSamples        As Long
    GetOutputSetting            As Long
    SetOutputSetting            As Long
    GetOutputCount              As Long
    GetOutputProps              As Long
    SetOutputProps              As Long
    GetOutputFormatCount        As Long
    GetOutputFormat             As Long
    GetOutputNumberForStream    As Long
    GetStreamNumberForOutput    As Long
    GetMaxOutputSampleSize      As Long
    GetMaxStreamSampleSize      As Long
    OpenStream                  As Long
End Type

Public Type WM_MEDIA_TYPE
    majortype                   As GUID
    subtype                     As GUID
    bFixedSizeSamples           As Long
    bTemporalCompression        As Long
    lSampleSize                 As Long
    formattype                  As GUID
    pUnk                        As Long
    cbFormat                    As Long
    pbFormat                    As Long
End Type


Public Const WMMEDIATYPE_Audio      As String _
    = "{73647561-0000-0010-8000-00AA00389B71}"

Public Const WMMEDIASUBTYPE_PCM     As String _
    = "{00000001-0000-0010-8000-00AA00389B71}"


Public Const WMFORMAT_WaveFormatEx  As String _
    = "{05589f81-c356-11ce-bf01-00aa0055595a}"


Public Const IID_IWMHeaderInfo      As String _
    = "{96406BDA-2B2B-11d3-B36B-00C04F6108FF}"

Public Const IID_IWMHeaderInfo2     As String _
    = "{15CF9781-454E-482e-B393-85FAE487A810}"

Public Const IID_IWMHeaderInfo3     As String _
    = "{15CC68E3-27CC-4ecd-B222-3F5D02D80BD5}"


Public Const IID_IWMProfileManager  As String _
    = "{d16679f2-6ca0-472d-8d31-2f5d55aee155}"

Private Const IID_IWMProfileManager2 As String _
    = "{7A924E51-73C1-494d-8019-23D37ED9B89A}"


Public Const IID_IWMCodecInfo       As String _
    = "{A970F41E-34DE-4a98-B3BA-E4B3CA7528F0}"

Public Const IID_IWMCodecInfo2     As String _
    = "{AA65E273-B686-4056-91EC-DD768D4DF710}"


Public Const IID_IWMMediaProps      As String _
    = "{96406BCE-2B2B-11d3-B36B-00C04F6108FF}"


Public Const WMMEDIASUBTYPE_DRM     As String _
    = "{00000009-0000-0010-8000-00AA00389B71}"



Public Const IID_IUnknown           As String _
    = "{00000000-0000-0000-C000-000000000046}"

Public Const IID_IWMOutputMediaProps As String _
    = "{96406BD7-2B2B-11d3-B36B-00C04F6108FF}"

Public Const IID_IWMSyncReader      As String _
    = "{9397F121-7705-4dc9-B049-98B698188414}"

Public Const IID_INSSBuffer         As String _
    = "{E1CD3524-03D7-11d2-9EED-006097D2D7CF}"


Public Const attr_WMDuration    As String = "Duration"
Public Const attr_WMBitrate     As String = "Bitrate"
Public Const attr_WMSeekable    As String = "Seekable"
Public Const attr_WMStridable   As String = "Stridable"
Public Const attr_WMBroadcast   As String = "Broadcast"
Public Const attr_WMProtected   As String = "Is_Protected"
Public Const attr_WMTrusted     As String = "Is_Trusted"
Public Const attr_WMSigName     As String = "Signature_Name"
Public Const attr_WMHasAudio    As String = "HasAudio"
Public Const attr_WMHasImage    As String = "HasImage"
Public Const attr_WMHasScript   As String = "HasScript"
Public Const attr_WMHasVideo    As String = "HasVideo"
Public Const attr_WMCurBitrate  As String = "CurrentBitrate"
Public Const attr_WMOptBitrate  As String = "OptimalBitrate"
Public Const attr_WMSkipBackw   As String = "Can_Skip_Backward"
Public Const attr_WMSkipForw    As String = "Can_Skip_Forward"
Public Const attr_WMNumFrames   As String = "NumberOfFrames"
Public Const attr_WMFileSize    As String = "FileSize"

Public Const attr_WMTitle       As String = "Title"
Public Const attr_WMAuthor      As String = "Author"
Public Const attr_WMDescript    As String = "Description"
Public Const attr_WMRating      As String = "Rating"
Public Const attr_WMCopyright   As String = "Copyright"

Public Const attr_WMAlbumTitle  As String = "WM/AlbumTitle"
Public Const attr_WMAlbumArtist As String = "WM/AlbumArtist"
Public Const attr_WMTrack       As String = "WM/Track"
Public Const attr_WMGenre       As String = "WM/Genre"
Public Const attr_WMYear        As String = "WM/Year"
Public Const attr_WMGenreID     As String = "WM/GenreID"
Public Const attr_WMMCDI        As String = "WM/MCDI"
Public Const attr_WMComposer    As String = "WM/Composer"
Public Const attr_WMLyrics      As String = "WM/Lyrics"
Public Const attr_WMTrackNumber As String = "WM/TrackNumber"
Public Const attr_WMIsVBR       As String = "IsVBR"
