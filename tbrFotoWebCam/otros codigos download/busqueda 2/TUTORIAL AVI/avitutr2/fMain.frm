VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "AVIFile Tutorial Stream Info"
   ClientHeight    =   1305
   ClientLeft      =   1800
   ClientTop       =   1965
   ClientWidth     =   4185
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1305
   ScaleWidth      =   4185
   Begin VB.CommandButton cmdOpenAVIFile 
      Caption         =   "cmdOpenAVIFile"
      Height          =   615
      Left            =   1305
      TabIndex        =   0
      Top             =   345
      Width           =   1575
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call AVIFileInit   '// opens AVIFile library
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call AVIFileExit   '// releases AVIFile library
End Sub

Private Sub cmdOpenAVIFile_Click()
    Dim res As Long         'result code
    Dim ofd As cFileDlg     'OpenFileDialog class
    Dim szFile As String    'filename
    Dim pAVIFile As Long    'pointer to AVI file interface (PAVIFILE handle)
    Dim pAVIStream As Long  'pointer to AVI stream interface (PAVISTREAM handle)
    Dim numFrames As Long   'number of frames in video stream
    Dim firstFrame As Long  'position of the first video frame
    Dim fileInfo As AVI_FILE_INFO       'file info struct
    Dim streamInfo As AVI_STREAM_INFO   'stream info struct
    
    'Get the name of an AVI file to work with
    Set ofd = New cFileDlg
    With ofd
        .OwnerHwnd = Me.hWnd
        .Filter = "AVI Files|*.avi"
        .DlgTitle = "Open AVI File"
    End With
    res = ofd.VBGetOpenFileNamePreview(szFile)
    If res = False Then GoTo ErrorOut
    
    'Open the AVI File and get a file interface pointer (PAVIFILE)
    res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo ErrorOut
 
    'Get the first available video stream (PAVISTREAM)
    res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If res <> AVIERR_OK Then GoTo ErrorOut
    
    'get the starting position of the stream (some streams may not start simultaneously)
    firstFrame = AVIStreamStart(pAVIStream)
    If firstFrame = -1 Then GoTo ErrorOut 'this function returns -1 on error
    
    'get the length of video stream in frames
    numFrames = AVIStreamLength(pAVIStream)
    If numFrames = -1 Then GoTo ErrorOut ' this function returns -1 on error
    
    MsgBox "PAVISTREAM handle is " & pAVIStream & vbCrLf & _
            "Video stream length - " & numFrames & vbCrLf & _
            "Stream starts on frame #" & firstFrame & vbCrLf & _
            "File and Stream info will be written to Immediate Window (from IDE - Ctrl+G to view)", vbInformation, App.title
    
    'get file info struct (UDT)
    res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut
    
    'print file info to Debug Window
    Call DebugPrintAVIFileInfo(fileInfo)
    
    'get stream info struct (UDT)
    res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut
    
    'print stream info to Debug Window
    Call DebugPrintAVIStreamInfo(streamInfo)

ErrorOut:
    If pAVIStream <> 0 Then
        Call AVIStreamRelease(pAVIStream) '//closes video stream
    End If
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile) '// closes the file
    End If
    
    If (res <> AVIERR_OK) Then 'if there was an error then show feedback to user
        MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.title
    End If
End Sub
