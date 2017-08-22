VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVisualization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Visualization Class Module by Vesa Piittinen aka Merri
' http://merri.net
'
' this class module gives you data out of soundcard required for visualizations


Option Explicit


Private Type WaveFormat
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type Wave
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type


Private Const WAVE_FORMAT_PCM = 1
Private Const WHDR_DONE = &H1&


'Visualization class API declarations
Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal Err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long



Private AudioData(511) As Integer               'for holding audio data
Private Device As Long                          'for holding device
Private ErrorsList As New Collection            'for storing the errors
Private PCM As Wave                             'for receiving audio data

Public IsReady As Boolean                       'is device ready


Private Sub Class_Initialize()
    'try opening device
    InitDevice
End Sub
Private Sub Class_Terminate()
    'close device
    CloseDevice
End Sub


'Visualization class error handling
Private Sub AddError(ByVal Where As String, ByVal Number As Long, ByVal Description As String)
    'add the error always as the first item
    If ErrorsList.Count Then
        ErrorsList.Add Where & "|" & Number & "|" & Description, , 1
    Else
        ErrorsList.Add Where & "|" & Number & "|" & Description
    End If
End Sub
Private Function CheckError(ByVal Where As String, ByVal ErrorNumber As Long) As Boolean
    Dim ErrorString As String * 128
    'returns true if no error
    CheckError = (ErrorNumber = 0)
    'get the error if occurred
    If Not CheckError Then
        'get error string
        waveInGetErrorText ErrorNumber, ErrorString, Len(ErrorString)
        'store error
        AddError Where, ErrorNumber, ErrorString
    End If
End Function
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


'Visualization class initialization, reading and closing
Private Sub CloseDevice()
    On Error Resume Next
    'close recording device
    'if you want to add error detection, replace Call with CheckError "CloseDevice",
    Call waveInReset(Device)
    Call waveInStop(Device)
    Call waveInClose(Device)
    Device = 0
    IsReady = False
    'report any error which might have occurred
    If Err Then AddError "CloseDevice", Err.Number, Err.Description
End Sub
Private Sub InitDevice()
    Dim WAV As WaveFormat
    On Error GoTo ErrorHandler
    'make sure device is closed
    CloseDevice
    With WAV
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2
        .SamplesPerSec = 44100
        .BitsPerSample = 16
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    'make device ready
    If CheckError("InitDevice", waveInOpen(Device, 0&, VarPtr(WAV), 0&, 0&, 0&)) Then
        'start recording
        If CheckError("InitDevice", waveInStart(Device)) Then
            With PCM
                'this links AudioData to contain information read using PCM
                .lpData = VarPtr(AudioData(0))
                .dwBufferLength = 1024
                .dwFlags = 0
            End With
            'success!
            IsReady = True
        End If
    End If
    Exit Sub
ErrorHandler:
    AddError "InitDevice", Err.Number, Err.Description
    'close device
    CloseDevice
End Sub


'Visualization class
Public Function GetData(ByVal Index As Integer) As Integer
    'error correction
    If Index < 0 Then Index = 0
    If Index > 511 Then Index = 511
    'return data
    GetData = AudioData(Index)
End Function
Public Sub Update()
    On Error GoTo ErrorHandler
    'read buffer: PCM is linked to AudioData and thus AudioData holds information
    If CheckError("Update", waveInPrepareHeader(Device, VarPtr(PCM), Len(PCM))) Then
        'if you want to add error detections, replace Call with CheckError "Update",
        'if you add it, you'll eventually get error# 33 sometimes
        Call waveInAddBuffer(Device, VarPtr(PCM), Len(PCM))
        'if wave buffer has enough for us to process, get it
        If (PCM.dwFlags And WHDR_DONE) = WHDR_DONE Then Call waveInUnprepareHeader(Device, VarPtr(PCM), Len(PCM))
    End If
    'success!
    Exit Sub
ErrorHandler:
    AddError "Update", Err.Number, Err.Description
    'close device
    CloseDevice
End Sub
