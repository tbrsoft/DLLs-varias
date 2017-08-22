Attribute VB_Name = "ActiveMovieControl"
'*******************************************************************************
'***            ActiveMove Control Module   -  By Fade
'***
'***    This is an ActiveMovie object usage Module
'***    Programmed based on the Microsoft DXSDK Code Samples.
'***    So there is more there for more advanced options :)
'***
'***    This module basicly handeles viewing of Video Files in most formats
'***    easy to use, just follow the basic methods, and remember to unload
'***    the objects using the 'Unload' method if changing mine.
'***
'***    OH! don't forget to set a reference to the 'ActiveMovie Control type library'
'***    in your project or all of this won't work :)))
'*******************************************************************************
Option Explicit
Option Base 0
Option Compare Text

Private m_dblRate As Double                'Rate in Frames Per Second
Private m_bstrFileName As String           'Loaded Filename
Private m_dblRunLength As Double           'Duration in seconds
Private m_dblStartPosition As Double       'Start position in seconds
Public m_boolVideoRunning As Boolean       'Flag used to trigger clock

Private dblPosition As Double ' Current Play position

Private m_objBasicAudio  As IBasicAudio      'Basic Audio Object
Private m_objBasicVideo As IBasicVideo       'Basic Video Object
Private m_objMediaEvent As IMediaEvent       'MediaEvent Object
Private m_objVideoWindow As IVideoWindow     'VideoWindow Object
Private m_objMediaControl As IMediaControl   'MediaControl Object
Private m_objMediaPosition As IMediaPosition 'MediaPosition Object
            
            
                      
    ' ****************************************************
    ' ****   Main Video Loading method
    ' ****      Use this method to load video file
Sub RunVideoContent(ByVal path As String, Optional ByVal DontMaintainRatio As Boolean, Optional ByVal FullScreen As Boolean)
    Dim nCount As Long
    Dim sScale As Double
    Dim topMod As Long
    On Local Error GoTo ErrLine
     
        ' NOTE: to get the clip duration use - m_dblRunLength

        ' Initialize global variables based on the
        ' contents of the file:
        '   m_bstrFileName - name of file name selected by the user
        '   m_dblRunLength = length of the file; duration
        '   m_dblStartPosition - point at which to start playing clip
        '   m_objMediaControl, m_objMediaEvent, m_objMediaPosition,
        '   m_objBasicAudio, m_objVideoWindow - programmable objects
    
        'clean up memory (in case a file was previously opened)
    UnloadActiveMovieControl
    
        ' Setting file to object
    m_bstrFileName = path
    
        'Instantiate a filter graph for the requested
        'file format.
    Set m_objMediaControl = New FilgraphManager
    Call m_objMediaControl.RenderFile(m_bstrFileName)
    
        'Setup the IBasicAudio object (this
        'is equivalent to calling QueryInterface()
        'on IFilterGraphManager). Initialize the volume
        'to the maximum value.
    
        ' Some filter graphs don't render audio
        ' In this sample, skip setting volume property
    Set m_objBasicAudio = m_objMediaControl
    m_objBasicAudio.Volume = 0
    m_objBasicAudio.Balance = 0
    
        'Setup the IVideoWindow object. Remove the
        'caption, border, dialog frame, and scrollbars
        'from the default window. Position the window.
        'Set the parent to the app's form.
    Set m_objVideoWindow = m_objMediaControl
    m_objVideoWindow.WindowStyle = CLng(&H6000000)
    m_objVideoWindow.Left = 0
        ' Getting Scale Ratio
    sScale = m_objVideoWindow.Height / m_objVideoWindow.Width
        ' Setting object width
    m_objVideoWindow.Width = Video_ActiveMovie.Video.Width
    If Not (DontMaintainRatio) Then
        m_objVideoWindow.Height = Video_ActiveMovie.Video.Width * sScale
        topMod = (Video_ActiveMovie.Video.Height - m_objVideoWindow.Height) / 2
    Else
        m_objVideoWindow.Height = Video_ActiveMovie.Video.Height
    End If
    m_objVideoWindow.Top = topMod
        ' Setting FullScreen Mode
    m_objVideoWindow.FullScreenMode = FullScreen
        'reset the video window owner - The surface the video is implemented upon
    m_objVideoWindow.Owner = Video_ActiveMovie.Video.hWnd
    
        'Setup the IMediaEvent object for the
        'sample toolbar (run, pause, play).
    Set m_objMediaEvent = m_objMediaControl
    
        'Setup the IMediaPosition object so that we
        'can display the duration of the selected
        'video as well as the elapsed time.
    Set m_objMediaPosition = m_objMediaControl
    
        'set the playback rate given the desired optional
    m_objMediaPosition.Rate = 1 ' Normal play rate
                                ' NOTE: you can set values like 1.5 for 150% speed, pretty nice
        ' Use user-established playback rate
    m_dblRate = m_objMediaPosition.Rate
        ' getting play length
    m_dblRunLength = Round(m_objMediaPosition.Duration, 2)
        ' Reset start position to 0
    m_dblStartPosition = 0
    
        ' Play the file
    PlayActiveMovie
    Exit Sub
    
ErrLine:
    Err.Clear
    Resume Next
End Sub
            
    ' ****************************************************
    ' ****   Unloading Control from memory
    ' ****
Sub UnloadActiveMovieControl()
    On Local Error GoTo ErrLine
    
    'stop playback
    m_boolVideoRunning = False
    DoEvents
    'cleanup media control
    If Not m_objMediaControl Is Nothing Then
        m_objMediaControl.Stop
    End If
    'clean-up video window
    If Not m_objVideoWindow Is Nothing Then
        m_objVideoWindow.Left = Screen.Width * 8
        m_objVideoWindow.Height = Screen.Height * 8
        m_objVideoWindow.Owner = 0          'sets the Owner to NULL
    End If
            
    'clean-up & dereference
    If Not m_objBasicAudio Is Nothing Then Set m_objBasicAudio = Nothing
    If Not m_objBasicVideo Is Nothing Then Set m_objBasicVideo = Nothing
    If Not m_objMediaControl Is Nothing Then Set m_objMediaControl = Nothing
    If Not m_objVideoWindow Is Nothing Then Set m_objVideoWindow = Nothing
    If Not m_objMediaPosition Is Nothing Then Set m_objMediaPosition = Nothing
    Exit Sub
            
ErrLine:
    Err.Clear
End Sub
            

    ' ****************************************************
    ' ****   Control Methods
    ' ****      Play,Pause & Stop

Sub PlayActiveMovie()
    On Local Error GoTo errHandle
    
    'Invoke the MediaControl Run() method
    'and pause the video that is being
    'displayed through the predefined
    'filter graph.
    
    'Assign specified starting position dependent on state
    If CLng(m_objMediaPosition.CurrentPosition) < CLng(m_dblStartPosition) Then
        m_objMediaPosition.CurrentPosition = m_dblStartPosition
    ElseIf CLng(m_objMediaPosition.CurrentPosition) = CLng(m_dblRunLength) Then
        m_objMediaPosition.CurrentPosition = m_dblStartPosition
    End If
    
    m_boolVideoRunning = True
    Call m_objMediaControl.Run
    
    
    Exit Sub
errHandle:
    Err.Clear
    Resume Next
    'logerror
End Sub

Sub PauseActiveMovie()
    On Local Error GoTo errHandle
        ' Validating state
    If Not (m_boolVideoRunning) Then Exit Sub
        ' Pausing
    Call m_objMediaControl.Pause
        ' setting state
    m_boolVideoRunning = False
    
    Exit Sub
errHandle:
    Err.Clear
    'logerror
End Sub

Sub StopActiveMovie()
    On Local Error GoTo errHandle
        ' Validating state
    If Not (m_boolVideoRunning) Then Exit Sub
        ' Stopping
    Call m_objMediaControl.Stop
        ' setting state
    m_boolVideoRunning = False
        ' reset to the beginning of the video
    m_objMediaPosition.CurrentPosition = 0
    
    Exit Sub
errHandle:
    Err.Clear
    'logerror
End Sub
                        

    ' ****************************************************
    ' ****   Various Setting methods
    ' ****
Sub SetActiveMovieBalance(ByVal Value As Long)
    On Local Error GoTo ErrLine
    'Set the balance using the slider
    If Not m_objMediaControl Is Nothing Then _
        m_objBasicAudio.Balance = Value
    Exit Sub
ErrLine:
    Err.Clear
End Sub

Sub SetActiveMovieVolume(ByVal Value As Long)
    On Local Error GoTo ErrLine
            
    'Set the volume using the slider
    If Not m_objMediaControl Is Nothing Then _
        m_objBasicAudio.Volume = Value
    Exit Sub
        
ErrLine:
    Err.Clear
End Sub
            
    ' ****************************************************
    ' ****   Info retrival Functions
    ' ****
Function GetVideoLength() As Double
    GetVideoLength = m_dblRunLength
End Function
    
Function GetVideoPos() As Double
    dblPosition = m_objMediaPosition.CurrentPosition
    GetVideoPos = dblPosition
End Function

Function VideoRunning() As Boolean
    VideoRunning = m_boolVideoRunning
End Function
            
    ' ****************************************************
    ' ****   The Timer event
    ' ****
Public Sub ActiveMovieTimerEvent()
    Dim nReturnCode As Long
    
    On Local Error GoTo errHandle
    
    If m_boolVideoRunning = True Then
            'obtain return code
        Call m_objMediaEvent.WaitForCompletion(100, nReturnCode)
                
               
        If nReturnCode = 0 Then ' Playing
            
            'get the current position for display
            dblPosition = m_objMediaPosition.CurrentPosition
            
        Else    ' Stopped
                ' NOTE: only occurs when clip FINISHES playin
                ' Set State
            m_boolVideoRunning = False
                ' Send event
            Video_ActiveMovie.VideoFinishedEvent
        End If
    End If
    Exit Sub
errHandle:
    'NOTE: Keep this as this method repeatedly raises errors
    '      without caring for my mental health
    Err.Clear
    Resume Next
End Sub
            

