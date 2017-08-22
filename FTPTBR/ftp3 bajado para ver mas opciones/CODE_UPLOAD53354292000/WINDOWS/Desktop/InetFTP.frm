VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Angelfireftp 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Angelfire FTP"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8280
   Icon            =   "InetFTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txbPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   840
      Width           =   3105
   End
   Begin VB.TextBox txbUserName 
      Height          =   315
      Left            =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   3105
   End
   Begin VB.TextBox txbURL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Text            =   "ftp://ftp.angelfire.com"
      Top             =   120
      Width           =   3105
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1845
      Left            =   2280
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3495
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1665
      Left            =   2280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3495
   End
   Begin VB.ListBox lisServerFiles 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   2985
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2145
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2400
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Disconnect logon "
      Top             =   3360
      Width           =   2325
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2085
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Exit program"
      Top             =   4080
      Width           =   2325
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Connect to server"
      Top             =   1200
      Width           =   2325
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Upload selected local file to selected server dir"
      Top             =   2640
      Width           =   2325
   End
   Begin VB.CommandButton cmdDownLoad 
      Caption         =   "Download "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Download file from server to local"
      Top             =   1920
      Width           =   2325
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Password"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "User Name"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Angelfire Server"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   5655
   End
End
Attribute VB_Name = "Angelfireftp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' InetFTP.frm
'
' By Herman Liu
'
' Notes:
' (1) An FTP client, for transfer of files to and from server and client, text or binary.
'     Microsoft FTP site is defaulted for ready testing and use by readers. You may even
'     try to download a couple of files before browsing this code (To avoid overwriting,
'     program will first see if there is already in existence a file of same name, either
'     in download or in upload).
' (2) There are Help buttons to guide your use of this program.
' (3) File size will be reported before a downloading/uploading action.
' (4) VB6 only (VB5 Inet has many known bugs which require fixes from Service Pack).

Option Explicit

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags _
      As Long, ByVal dwReserved As Long) As Long


Const defaultURL = "ftp://ftp.microsoft.com"
Const defaultUserName = "yourusername"
Const defaultPassword = "yourpassword"
Const defaultEMailAddress = "youremailaddress"

Dim ConnectedFlag As Boolean
Dim ServerDirFlag As Boolean
Dim DownloadFlag As Boolean
Dim UploadFlag As Boolean
Dim FileSizeFlag As Boolean

Dim homeLen As Integer
Dim LocFilespec As String
Dim SerFilespec As String
Dim gFileSize As String




Private Sub Form_Load()
    GetStartingDefaults
    ConnectedFlag = False
    ClearFlags
    UpdButtons
End Sub


Private Sub GetStartingDefaults()
    txbURL.Text = defaultURL
    txbUserName.Text = ""
    txbPassword.Text = ""
End Sub



Private Sub cmdConnect_click()
     On Error Resume Next
     Dim tmp As String
     Dim i As Integer
     
     Inet1.Cancel
     Inet1.Execute , "CLOSE"
    
     Err.Clear
     On Error GoTo errHandler
    
     ClearFlags
    
     If Len(txbURL) < 6 Then
          MsgBox "No URL yet"
          Exit Sub
     End If
    
     If UCase(Left(txbURL, 6)) <> "FTP://" Then
          MsgBox "No FTP protocol entered in URL"
          Exit Sub
     End If
    
     lblStatus.Caption = "To connect ...."
    
       ' (Note we use txtURL.Text here; you can just use txtURL if you wish)
     Inet1.AccessType = icUseDefault
     Inet1.URL = LTrim(Trim(txbURL.Text))
     Inet1.UserName = LTrim(Trim(txbUserName.Text))
     Inet1.Password = LTrim(Trim(txbPassword.Text))
     Inet1.RequestTimeout = 40
            
       ' Will force to bring up Dialup Dialog if not already having a line
     ServerDirFlag = True
     Inet1.Execute , "DIR"
     Do While Inet1.StillExecuting
          DoEvents
          ' Connection not established yet, hence cannot
          ' try to fall back on ConnectedFlag to exit
     Loop
     txbURL.Text = Inet1.URL
     
          ' Home portion
     For i = 7 To Len(txbURL.Text)
          tmp = Mid(txbURL.Text, i, 1)
          If tmp = "/" Then
               Exit For
          End If
     Next i
     homeLen = i - 1
     
     If IsNetConnected() Then
          ConnectedFlag = True
          UpdButtons
     Else
          GoTo errHandler
     End If
     Exit Sub
    
errHandler:
    If icExecuting Then
           ' We place this here in case command for "CLOSE" failed.
           ' With Inet, one can never tell.
         If ConnectedFlag = False Then
              Exit Sub
         End If
        
         If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
              Inet1.Cancel
              If Inet1.StillExecuting Then
                  lblStatus.Caption = "System failed to cancel job"
              End If
         Else
              Resume
         End If
     End If
     ErrMsgProc "cmdConnect_Click"
End Sub



Private Sub cmdExit_Click()
    On Error Resume Next
    Inet1.Execute , "CLOSE"
    Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Inet1.Execute , "CLOSE"
    Unload Me
End Sub



Private Sub cmdPublic_Click()
    txbUserName.Text = "anonymous"
    txbPassword.PasswordChar = ""
    txbPassword.Text = defaultEMailAddress
    txbURL.SetFocus
End Sub



Private Sub cmdPrivate_Click()
    txbUserName.Text = defaultUserName
    txbPassword.PasswordChar = "*"
    txbPassword.Text = defaultPassword
    txbURL.SetFocus
End Sub



Private Sub cmdNil_Click()
    txbUserName.Text = ""
    txbPassword.Text = ""
    txbURL.SetFocus
End Sub




Private Sub cmdDisconnect_Click()
    On Error Resume Next
    Inet1.Cancel
    Inet1.Execute , "CLOSE"
    lblStatus.Caption = "Unconnected"
       ' Put back starting default
    GetStartingDefaults
    ConnectedFlag = False
    lisServerFiles.Clear
    ClearFlags
    UpdButtons
End Sub




Private Sub ClearFlags()
    ServerDirFlag = False
    DownloadFlag = False
    UploadFlag = False
    FileSizeFlag = False
End Sub




Private Sub UpdButtons()
    cmdConnect.Enabled = False
    cmdDownLoad.Enabled = False
    cmdUpload.Enabled = False
    cmdDisconnect.Enabled = False
    If ConnectedFlag Then
           ' Once connected, no interference to txbURL.text
         txbURL.Locked = True
         cmdDownLoad.Enabled = True
         cmdUpload.Enabled = True
         cmdDisconnect.Enabled = True
    Else
         txbURL.Locked = False
         cmdConnect.Enabled = True
    End If
End Sub




Private Sub cmdDownLoad_Click()
     On Error GoTo errHandler
     
     If ConnectedFlag = False Then
          MsgBox "No connection yet"
          Exit Sub
     ElseIf lisServerFiles.ListCount = 0 Then
          MsgBox "No server file listed yet"
          Exit Sub
     ElseIf Right(lisServerFiles.Text, 1) = "/" Then
          MsgBox "Selected item is a directory only." & vbCrLf & vbCrLf & _
             "To list files under that dir, double click on it."
          Exit Sub
     End If
    
     lblStatus.Caption = "Retreiving file..."
     SerFilespec = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & _
               "/" & lisServerFiles.Text
     SerFilespec = Right(SerFilespec, Len(SerFilespec) - 1)
     
        ' Use same file name and store it in current dir of local. Parse
        ' above SerFilespec and take only the file name as LocFileSpec.
     LocFilespec = SerFilespec
     Do While InStr(LocFilespec, "/") <> 0
         LocFilespec = Right(LocFilespec, Len(LocFilespec) - _
              InStr(LocFilespec, "/"))
     Loop
     
     If IsFileThere(LocFilespec) Then
          If MsgBox(LocFilespec & " already exist. Overwrite?", _
               vbYesNo + vbQuestion) = vbNo Then
               Exit Sub
          End If
     End If
     
     lblStatus.Caption = "Requesting for file size..."
     
     gFileSize = ""
     FileSizeFlag = True
     Inet1.Execute , "SIZE " & SerFilespec
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
         
     If gFileSize = "" Then
          MsgBox "Selected file has 0 byte content."
          Exit Sub
     Else
          If MsgBox("File size is " & gFileSize & " bytes." & vbCrLf & vbCrLf & _
                  "Proceed to download?", vbYesNo + vbQuestion) = vbNo Then
              Exit Sub
          End If
     End If
     
     DownloadFlag = True
     Inet1.Execute , "Get " & SerFilespec & " " & LocFilespec
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop

     lblStatus.Caption = "Connected"
     File1.Refresh
     Exit Sub
     
errHandler:
    If icExecuting Then
        If ConnectedFlag = False Then
            Exit Sub
        End If
        
        If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
            Inet1.Cancel
            If Inet1.StillExecuting Then
                lblStatus.Caption = "System failed to cancel job"
            End If
        Else
            Resume
        End If
    End If
    ErrMsgProc "cmdDownLoad_Click"
End Sub




' Assuming you have the appropriate privileges on the server
Private Sub cmdUpLoad_Click()
     On Error GoTo errHandler
     Dim tmpPath As String
     Dim tmpFile As String
     Dim bExist As Boolean
     Dim lFileSize As Long
     Dim i
     
     If ConnectedFlag = False Then
          MsgBox "No connection yet"
          Exit Sub
     ElseIf File1.ListCount = 0 Then
          MsgBox "No local file in current dir yet"
          Exit Sub
     ElseIf Not (Right(lisServerFiles.Text, 1) = "/") Then
          MsgBox "Selected server file item is not a directory"
          Exit Sub
     ElseIf lisServerFiles.Text = "../" Then
          MsgBox "No directory name selected yet"
          Exit Sub
     End If
    
     LocFilespec = tmpPath & File1.List(File1.ListIndex)
     If LocFilespec = "" Then
          MsgBox "No local file selected yet"
          Exit Sub
     End If
     
     lFileSize = FileLen(LocFilespec)
     If MsgBox("File size is " & CStr(lFileSize) & " bytes." & vbCrLf & vbCrLf & _
                  "Proceed to upload?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
     End If
    
     lblStatus.Caption = "Uploading file..."
     
     If Right(Dir1.Path, 1) <> "\" Then
          tmpPath = Dir1.Path & "\"
     Else
          tmpPath = Dir1.Path                   ' e.g. root "C:\"
     End If
     
     SerFilespec = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & _
               "/" & lisServerFiles.Text
          ' Remove the front "/" from above
     SerFilespec = Right(SerFilespec, Len(SerFilespec) - 1)
     
     SerFilespec = SerFilespec & File1.List(File1.ListIndex)
     
          ' In order to test whether same file on server already exists
     lblStatus.Caption = "Verifying existence of file of same name..."
     tmpPath = SerFilespec
     ServerDirFlag = True
     Inet1.Execute , "DIR " & tmpPath & "/*.*"
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
         
     bExist = False
     If lisServerFiles.ListCount > 0 Then
          For i = 0 To lisServerFiles.ListCount - 1
               tmpFile = lisServerFiles.List(i)
               If tmpFile = File1.List(File1.ListIndex) Then
                    bExist = True
                    Exit For
               End If
          Next i
     End If
         
          ' Go back
     ServerDirFlag = True
     Inet1.Execute , "DIR ../*"
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
          
          
     If bExist Then
          If MsgBox("File already exist in selected server dir.  Supersede?", _
                  vbYesNo + vbQuestion) = vbNo Then
               Exit Sub
          End If
     End If
     
     Exit Sub
         
     UploadFlag = True
     Inet1.Execute , "PUT " & LocFilespec & " " & SerFilespec
     
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop

     lblStatus.Caption = "Connected"
     Exit Sub
    
errHandler:
     If icExecuting Then
         If ConnectedFlag = False Then
              Exit Sub
         End If
        
         If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion) = vbYes Then
              Inet1.Cancel
              If Inet1.StillExecuting Then
                   lblStatus.Caption = "System failed to cancel job"
              End If
         Else
              Resume
         End If
     End If
     ErrMsgProc "cmdUpload_Click"
End Sub



Private Sub lbServerFilesHelp_Click()
     MsgBox "Help:" & vbCrLf & vbCrLf & _
          "To change dir, double click a directory item on list." & vbCrLf & _
          "   (To go up one level, click the '../' item)" & vbCrLf & vbCrLf & _
          "To select a file for download, highlight it then" & vbCrLf & _
          "   click Download button (will report file size)." & vbCrLf & vbCrLf & _
          "To upload a local file, highlight a server dir first," & vbCrLf & _
          "   highlight a local file, then click Upload button." & vbCrLf & vbCrLf
End Sub



Private Sub lblLocalFilesHelp_Click()
     MsgBox "Help:" & vbCrLf & vbCrLf & _
          "To see file size of a local file, double click the" & vbCrLf & _
            "   local file item." & vbCrLf & vbCrLf & _
          "For other Help, refer Server Files." & vbCrLf & vbCrLf
End Sub

Private Sub lisServerFiles_dblClick()
     On Error GoTo errHandler
     
     If Not (Right(lisServerFiles.Text, 1) = "/") Then
          Exit Sub
     End If
     
     Dim tmpDir As String, tmp As String
     Dim i
     If Trim(lisServerFiles.Text) = "../" Then
          For i = Len(txbURL.Text) To 7 Step -1
               tmp = Mid(txbURL.Text, i, 1)
               If tmp = "/" Then
                    Exit For
               End If
          Next i
          If i = 7 Then
               MsgBox "No upper level of dir"
               Exit Sub
          End If
          txbURL.Text = Left(txbURL.Text, i - 1)
             ' Relative dir
          tmpDir = "../*"
     Else
          txbURL.Text = txbURL.Text & "/" & _
                   Left(lisServerFiles.Text, Len(lisServerFiles.Text) - 1)
          tmpDir = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & "/*"
     End If
     ServerDirFlag = True
     Inet1.Execute , "DIR " & tmpDir
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
     Exit Sub
    
errHandler:
    Select Case Err.Number
        Case icExecuting
             Resume
        Case Else
             ErrMsgProc "lisServerFiles_dblClick"
     End Select
End Sub



Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub



Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub



Private Sub File1_dblClick()
    If File1.ListCount = 0 Then
         Exit Sub
    End If
    Dim lFileSize As Long
    lFileSize = FileLen(File1.List(File1.ListIndex))
    MsgBox CStr(lFileSize) & " bytes"
End Sub




Private Sub Inet1_StateChanged(ByVal State As Integer)
    On Error Resume Next
    Select Case State
        Case icError                                      ' 11
            lblStatus = Inet1.ResponseCode & ": " & Inet1.ResponseInfo
            Inet1.Execute , "CLOSE"
            lblStatus.Caption = "Unconnected"
            lisServerFiles.Clear
            ConnectedFlag = False
            ServerDirFlag = False
            DownloadFlag = False
            UpdButtons
            
        Case icResponseCompleted                          ' 12
            Dim bDone As Boolean
            Dim tmpData As Variant       ' GetChunk returns Variant type
            
            If ServerDirFlag = True Then
                 Dim dirData As String
                 Dim strEntry As String
                 Dim i As Integer, k As Integer
            
                 tmpData = Inet1.GetChunk(4096, icString)
                 dirData = dirData & tmpData
            
                 If dirData <> "" Then
                     lisServerFiles.Clear
                       ' Use relative address to allow one dir level up
                     lisServerFiles.AddItem ("../")
                     For i = 1 To Len(dirData) - 1
                          k = InStr(i, dirData, vbCrLf)        ' We don't want CRLF
                          strEntry = Mid(dirData, i, k - i)
                          If Right(strEntry, 1) = "/" Then
                               strEntry = Left(strEntry, Len(strEntry) - 1) & "/"
                          End If
                          If Trim(strEntry) <> "" Then
                               lisServerFiles.AddItem strEntry
                          End If
                          i = k + 1
                          DoEvents
                     Next i
                     lisServerFiles.ListIndex = 0
                 End If
                 
                 ServerDirFlag = False
                 lblStatus.Caption = "Dir completed"
                 
            ElseIf DownloadFlag Then
                 Dim varData As Variant
                 
                 bDone = False

                 Open LocFilespec For Binary Access Write As #1
    
                   ' Get first chunk
                 tmpData = Inet1.GetChunk(10240, icByteArray)
                 DoEvents
                 If Len(tmpData) = 0 Then
                      bDone = True
                 End If
                 Do While Not bDone
                      varData = tmpData
                      Put #1, , varData
                      tmpData = Inet1.GetChunk(10240, icByteArray)
                      DoEvents
                      If ConnectedFlag = False Then
                           Exit Sub
                      End If
                      If Len(tmpData) = 0 Then
                            bDone = True
                      End If
                 Loop
                 Close #1
                 DownloadFlag = False
                 DoEvents
                 lblStatus.Caption = "Download completed"
                 DownloadFlag = False
                 MsgBox "Download completed:" & vbCrLf & vbCrLf & _
                     "File in current dir, named  " & LocFilespec
                 
            ElseIf UploadFlag Then
                 lblStatus.Caption = "Connected"
                 UploadFlag = False
                 MsgBox "Download completed: File in " & LocFilespec
                 
            ElseIf FileSizeFlag Then
                 Dim sizeData As String
            
                 tmpData = Inet1.GetChunk(1024, icString)
                 DoEvents
                 If Len(tmpData) > 0 Then
                      sizeData = sizeData & tmpData
                 End If
                 
                 gFileSize = sizeData
                 FileSizeFlag = False
                 
            Else
                 lblStatus.Caption = "Connected"
            End If
            
            
        Case icNone                                       ' 0
            lblStatus.Caption = "Unknown State Possible Error"
        Case icResolvingHost                              ' 1
            lblStatus.Caption = "Resolving Host..."
        Case icHostResolved                               ' 2
            lblStatus.Caption = "Host Resolved - IP Address Found"
        Case icConnecting                                 ' 3
            lblStatus.Caption = "Connecting..."
        Case icConnected                                  ' 4
            lblStatus.Caption = "Connected"
        Case icRequesting                                 ' 5
            lblStatus.Caption = "Sending Request..."
        Case icRequestSent                                ' 6
            lblStatus.Caption = "Request Sent"
        Case icReceivingResponse                          ' 7
            lblStatus = "Receiving Data..."
        Case icResponseReceived                           ' 8
            lblStatus = "Data Received"
        Case icDisconnecting                              ' 9
            lblStatus.Caption = "Disconnecting..."
        Case icDisconnected                               '10
            lblStatus = "Disconnected"
    End Select
End Sub



Function IsNetConnected() As Boolean
    IsNetConnected = InternetGetConnectedState(0, 0)
End Function
                  


Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub



Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim i
    i = FreeFile
    Open inFileSpec For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
End Function

