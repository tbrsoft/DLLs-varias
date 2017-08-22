VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "AVIFile Tutorial Framework Project"
   ClientHeight    =   1425
   ClientLeft      =   1800
   ClientTop       =   1965
   ClientWidth     =   4440
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   4440
   Begin VB.CommandButton cmdOpenAVIFile 
      Caption         =   "cmdOpenAVIFile"
      Height          =   615
      Left            =   1433
      TabIndex        =   0
      Top             =   405
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
    Dim pAVIFile As Long          'pointer to AVI File (PAVIFILE handle)
    
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
 
    MsgBox "PAVIFILE handle is " & pAVIFile, vbInformation, App.title
'//
'// Place functions here that interact with the open file.
'//

ErrorOut:
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile) '// closes the file
    End If
    
    If (res <> AVIERR_OK) Then 'if there was an error then show feedback to user
        MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.title
    End If
End Sub
