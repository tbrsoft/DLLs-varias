VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActiveMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ActiveMovie Class Module by Vesa Piittinen aka Merri
' http://merri.net
'
' this class requires the ActiveMovie control type library reference
' go Project > References to add it


Option Explicit


Private ErrorsList As New Collection            'for storing the errors

Private Owner As Long                           'for storing the video output owner information
Private Filename As String                      'for storing the filename

Private VidWindow As IVideoWindow               'ActiveMovie control for video output
Private MediaControl As IMediaControl           'ActiveMovie control for handling media
Private MediaPos As IMediaPosition              'ActiveMovie control for media position
Private IAudio  As IBasicAudio                  'ActiveMovie control for audio
Private IVideo As IBasicVideo                   'ActiveMovie control for video
Private IEvent As IMediaEvent                   'ActiveMovie control for media events


'ActiveMovie class
Private Sub Class_Terminate()
    'make sure we unload properly
    CloseFile
End Sub


'ActiveMovie class error handling
Private Sub AddError(ByVal Where As String, ByVal Number As Long, ByVal Description As String)
    'add the error always as the first item
    If ErrorsList.Count Then
        ErrorsList.Add Where & "|" & Number & "|" & Description, , 1
    Else
        ErrorsList.Add Where & "|" & Number & "|" & Description
    End If
End Sub
Public Sub ClearErrors()
    'remove all errors
    Do While ErrorsList.Count: ErrorsList.Remove 1: Loop
End Sub
Public Function Errors() As Integer
    'return the number of errors
    Errors = ErrorsList.Count
End Function
Public Function GetError(ByVal Index As Integer) As String
    'check for valid index
    If Index < 1 Or Index > ErrorsList.Count Then Exit Function
    'return error
    GetError = ErrorsList(Index)
End Function
Public Function GetLastError() As String
    'return last error
    If ErrorsList.Count Then GetLastError = ErrorsList(1)
End Function


'ActiveMovie class file handling
Public Function CloseFile() As Boolean
    On Error GoTo ErrorHandler
    'if we have nothing to close, quit
    If MediaControl Is Nothing Then Exit Function
    'stop
    MediaControl.Stop
    'free controls
    Set IAudio = Nothing
    Set IVideo = Nothing
    Set VidWindow = Nothing
    Set IEvent = Nothing
    Set MediaPos = Nothing
    Set MediaControl = Nothing
    'success!
    CloseFile = True
    Exit Function
ErrorHandler:
    AddError "CloseFile", Err.Number, Err.Description
End Function
Public Function OpenFile(ByVal hwndOwner As Long, ByVal File As String) As Boolean
    On Error GoTo ErrorHandler
    'check if already has a file
    If Not MediaControl Is Nothing Then CloseFile
    'prepare media
    Set MediaControl = New FilgraphManager
    'open file
    MediaControl.RenderFile File
    'to be able to control the file
    Set IAudio = MediaControl
    Set IVideo = MediaControl
    Set VidWindow = MediaControl
    Set IEvent = MediaControl
    Set MediaPos = MediaControl
    'remember filename
    Filename = File
    'set owner
    Owner = hwndOwner
    'success!
    OpenFile = True
    Exit Function
ErrorHandler:
    AddError "OpenFile", Err.Number, Err.Description
End Function


'ActiveMovie class playback handling
Public Function PauseFile() As Boolean
    On Error GoTo ErrorHandler
    'pause file
    MediaControl.Pause
    'success!
    PauseFile = True
    Exit Function
ErrorHandler:
    AddError "PauseFile", Err.Number, Err.Description
End Function
Public Function PlayFile() As Boolean
    On Error GoTo ErrorHandler
    'play file
    MediaControl.Run
    'initialize video output
    InitVideo
    'success!
    PlayFile = True
    Exit Function
ErrorHandler:
    AddError "PlayFile", Err.Number, Err.Description
End Function
Public Function StopFile() As Boolean
    On Error GoTo ErrorHandler
    'stop file
    MediaControl.Stop
    'success!
    StopFile = True
    Exit Function
ErrorHandler:
    AddError "StopFile", Err.Number, Err.Description
End Function


'ActiveMovie class playback position handling
Public Function Length() As Double
    If MediaPos Is Nothing Then Exit Function
    'return position
    Length = MediaPos.Duration
End Function
Public Function Position() As Double
    If MediaPos Is Nothing Then Exit Function
    'return position
    Position = MediaPos.CurrentPosition
End Function
Public Function Remaining() As Double
    If MediaPos Is Nothing Then Exit Function
    'return remaining
    Remaining = MediaPos.Duration - MediaPos.CurrentPosition
End Function
Public Sub SeekTo(ByVal NewPosition As Double)
    If MediaPos Is Nothing Then Exit Sub
    'set position
    MediaPos.CurrentPosition = NewPosition
End Sub



Public Function FormatHMS(ByVal Value As Double) As String
    Dim Hours As Integer, Minutes As Byte, Seconds As Byte
    On Error GoTo ErrorHandler
    'get hours
    Hours = Value \ 3600
    'get minutes
    Minutes = (Value Mod 3600) \ 60
    'get seconds
    Seconds = Value Mod 60
    'check for hours
    If CBool(Hours) Then
        'view with hours
        FormatHMS = Hours & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
    Else
        'view with minutes and seconds only
        FormatHMS = Minutes & ":" & Format(Seconds, "00")
    End If
    'success!
    Exit Function
ErrorHandler:
    AddError "FormatHMS", Err.Number, Err.Description
End Function
Public Function Move(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As Boolean
    On Error GoTo ErrorHandler
    'if there is no video output, quit
    If VidWindow Is Nothing Then Exit Function
    'initialize video
    InitVideo
    'resize video
    VidWindow.SetWindowPosition Left, Top, Width, Height
    'success!
    Move = True
    Exit Function
ErrorHandler:
    AddError "Move", Err.Number, Err.Description
End Function
Public Function VideoHeight() As Long
    On Error GoTo ErrorHandler
    If IVideo Is Nothing Then Exit Function
    'return video height
    VideoHeight = IVideo.SourceHeight
    Exit Function
ErrorHandler:
    AddError "VideoHeight", Err.Number, Err.Description
    VideoHeight = 1
End Function
Public Function VideoWidth() As Long
    On Error GoTo ErrorHandler
    If IVideo Is Nothing Then Exit Function
    'return video width
    VideoWidth = IVideo.SourceWidth
    Exit Function
ErrorHandler:
    AddError "VideoWidth", Err.Number, Err.Description
    VideoWidth = 1
End Function


'ActiveMovie class additional subs and functions
Private Sub InitVideo()
    'set video output for media
    Set VidWindow = MediaControl
    'make borderless
    VidWindow.WindowStyle = CLng(&H6000000)
    'set owner
    VidWindow.Owner = Owner
End Sub
