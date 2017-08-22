VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "AVIFile Tutorial Recompress"
   ClientHeight    =   1725
   ClientLeft      =   1800
   ClientTop       =   1965
   ClientWidth     =   4185
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   4185
   Begin VB.CommandButton cmdCancelSave 
      Caption         =   "cmdCancelSave"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2295
      TabIndex        =   2
      Top             =   315
      Width           =   1710
   End
   Begin VB.TextBox txtStatus 
      Enabled         =   0   'False
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Text            =   "Select AVI file to copy"
      Top             =   1170
      Width           =   3840
   End
   Begin VB.CommandButton cmdOpenAVIFile 
      Caption         =   "cmdOpenAVIFile"
      Height          =   615
      Left            =   180
      TabIndex        =   0
      Top             =   315
      Width           =   1710
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 'don't allow unload during write
 If cmdCancelSave.Enabled = True Then Cancel = True
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
    Dim szFileOut As String 'filename to create new AVI file
    Dim pAVIFileOut As Long 'pointer to AVI file interface for new AVI file
    Dim pAVIStreamOut As Long           'pointer to AVI stream for new file
    Dim opts As AVI_COMPRESS_OPTIONS    'compression options
    Dim pOpts As Long                   'temporary pointer to UDT to pass to API

    
    'Get the name of an AVI file to work with
    Set ofd = New cFileDlg
    With ofd
        .OwnerHwnd = Me.hWnd
        .Filter = "AVI Files|*.avi"
        .DlgTitle = "Choose AVI File to Copy Video From"
    End With
    res = ofd.VBGetOpenFileNamePreview(szFile)
    If res = False Then GoTo ErrorOut
    
    'Open the AVI File and get a file interface pointer (PAVIFILE)
    res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo ErrorOut
 
    'Get the first available video stream (PAVISTREAM)
    res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If res <> AVIERR_OK Then GoTo ErrorOut
    
    'Get the name of an AVI file to work with
    ofd.DlgTitle = "Choose Location and Name to Save New AVI File"
    ofd.DefaultExt = "avi"
    szFileOut = "MyFile.avi" 'suggested name to prompt users with
    res = ofd.VBGetSaveFileName(szFileOut)
    If res = False Then
        MsgBox "User cancelled - no file saved.", vbInformation, App.title
        GoTo ErrorOut
    End If
    
    DoEvents 'let screen redraw after showing save file dialog
    
    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT array (in this case a single UDT)
    pOpts = VarPtr(opts)
    res = AVISaveOptions(Me.hWnd, _
                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE Or ICMF_CHOOSE_PREVIEW, _
                        1, _
                        pAVIStream, _
                        pOpts)  'returns TRUE if User presses OK, FALSE if Cancel
    If res <> 1 Then 'This function returns True on success so must be handled differently (In C TRUE = 1)
        MsgBox "AviSaveOptions returned an error!", vbCritical, App.title
        res = 0 'don't show second MsgBox in error handler
        GoTo ErrorOut
    End If
    
    DoEvents 'let screen redraw after showing Options window
    
    'recompress the stream with user options
    res = AVIMakeCompressedStream(pAVIStreamOut, pAVIStream, opts, 0&)
    If res <> AVIERR_OK Then
        Call AVISaveOptionsFree(1, pOpts) 'avoid resource leak on error here by freeing comp options
        GoTo ErrorOut
    End If


    gfAbort = False 'make sure abort flag is reset
    cmdOpenAVIFile.Enabled = False 'prevent reentrancy
    cmdCancelSave.Enabled = True 'allow user to cancel
    'save the video stream to new filename
    'Careful! this API requires a pointer to a pointer to an array of UDTs (in this case a single UDT)
    'this function will call the Callback function in mAVIDecs.bas and show percent written
    'user can also abort the file write via this callback function
    pOpts = VarPtr(opts) 'make sure pointer is still valid
    res = AVISave(szFileOut, 0&, AddressOf AVISaveCallback, 1, pAVIStreamOut, pOpts)
    'if user cancelled save give feedback
    If res = AVIERR_USERABORT Then
        txtStatus = "User cancelled!"
    Else
        txtStatus = "Finished!"
    End If
    
    'Careful! this API requires a pointer to a pointer to an array of UDTs (in this case a single UDT)
    Call AVISaveOptionsFree(1, pOpts) 'free resources


ErrorOut:
    cmdCancelSave.Enabled = False
    cmdOpenAVIFile.Enabled = True
    
    If pAVIStream <> 0 Then
        Call AVIStreamRelease(pAVIStream) '//closes video stream
    End If
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile) '// closes the file
    End If
    
    If (res <> AVIERR_OK) Then 'if there was an error then show feedback to user
        If res <> AVIERR_USERABORT Then 'don't show msg if user aborted save
            MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.title
        End If
    End If
End Sub

Private Sub cmdCancelSave_Click()
    gfAbort = True 'cancel file save
    'see callback function in mAVIDecs.bas
End Sub
