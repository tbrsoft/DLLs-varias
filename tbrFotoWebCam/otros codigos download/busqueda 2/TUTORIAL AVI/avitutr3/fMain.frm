VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "AVIFile Tutorial AVI to BMP"
   ClientHeight    =   1860
   ClientLeft      =   1800
   ClientTop       =   1965
   ClientWidth     =   4410
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1860
   ScaleWidth      =   4410
   Begin VB.TextBox txtStatus 
      Enabled         =   0   'False
      Height          =   330
      Left            =   270
      TabIndex        =   1
      Text            =   "No AVI File Selected"
      Top             =   1125
      Width           =   3840
   End
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
    Dim dib As cDIB
    Dim pGetFrameObj As Long    'pointer to GetFrame interface
    Dim pDIB As Long            'pointer to packed DIB in memory
    Dim bih As BITMAPINFOHEADER 'infoheader to pass to GetFrame functions
    Dim i As Long
    
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
    
'    MsgBox "PAVISTREAM handle is " & pAVIStream & vbCrLf & _
'            "Video stream length - " & numFrames & vbCrLf & _
'            "Stream starts on frame #" & firstFrame & vbCrLf & _
'            "File and Stream info will be written to Immediate Window (from IDE - Ctrl+G to view)", vbInformation, App.title
'
    'get file info struct (UDT)
    res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut
    
'    'print file info to Debug Window
'    Call DebugPrintAVIFileInfo(fileInfo)
    
    'get stream info struct (UDT)
    res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut
    
'    'print stream info to Debug Window
'    Call DebugPrintAVIStreamInfo(streamInfo)
    
    'set bih attributes which we want GetFrame functions to return
    With bih
        .biBitCount = 24
        .biClrImportant = 0
        .biClrUsed = 0
        .biCompression = BI_RGB
        .biHeight = streamInfo.rcFrame.bottom - streamInfo.rcFrame.top
        .biPlanes = 1
        .biSize = 40
        .biWidth = streamInfo.rcFrame.right - streamInfo.rcFrame.left
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight 'calculate total size of RGBQUAD scanlines (DWORD aligned)
    End With
    
    'init AVISTreamGetFrame* functions and create GETFRAME object
    'pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, ByVal AVIGETFRAMEF_BESTDISPLAYFMT) 'tell AVIStream API what format we expect and input stream
    pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, bih) 'force function to return 24bit DIBS
    If pGetFrameObj = 0 Then
        MsgBox "No suitable decompressor found for this video stream!", vbInformation, App.title
        GoTo ErrorOut
    End If
    
    'create a DIB class to load the frames into
    Set dib = New cDIB
    For i = firstFrame To (numFrames - 1) + firstFrame
        pDIB = AVIStreamGetFrame(pGetFrameObj, i)  'returns "packed DIB"
        If dib.CreateFromPackedDIBPointer(pDIB) Then
            Call dib.WriteToFile(App.Path & "\" & i & ".bmp")
            txtStatus = "Bitmap " & i + 1 & " of " & numFrames & " written to app folder"
            txtStatus.Refresh
        Else
            
        End If
    Next
    
    Set dib = Nothing
    


ErrorOut:
    If pGetFrameObj <> 0 Then
        Call AVIStreamGetFrameClose(pGetFrameObj) '//deallocates the GetFrame resources and interface
    End If
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

