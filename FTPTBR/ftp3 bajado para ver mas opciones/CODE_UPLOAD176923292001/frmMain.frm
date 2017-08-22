VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AutoUpload"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   3480
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto Close"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2790
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin AutoUP.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   240
      TabIndex        =   2
      Top             =   2460
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   291
      Min             =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Abort Upload"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ListBox ServerLog 
      Height          =   2205
      ItemData        =   "frmMain.frx":0442
      Left            =   120
      List            =   "frmMain.frx":0444
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
   Begin MSWinsockLib.Winsock data 
      Left            =   2160
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin MSWinsockLib.Winsock con 
      Left            =   1680
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nr1 As Integer, Nr2 As Integer, num1 As String
Dim LocalIP As String
Dim prts As String
Dim bReplied As Boolean
Dim heard As Boolean
Dim lTime As Long
Function SetDownload()
Dim II As Integer
Dim numberof As Integer

numberof = 1

 ' will save to your current directory
 ChDir "c:\windows\desktop"
 local_dir = "c:\windows\desktop"
 
 
 con.SendData "TYPE I" + Chr(13) + Chr(10)
        
 WaitFor (200)
 
      For II = LBound(fName) To UBound(fName)
        Doing_Download = True
        open_file (ExtractName(fName(II)))
        client.File_Name = fName(II)
        Call ipeer
        con.SendData prts + Chr(13) + Chr(10)
        ServerLog.AddItem prts
        WaitFor (200)
        con.SendData "RETR " & ExtractName(fName(II)) + Chr(13) + Chr(10)
        WaitFor (150)
        ServerLog.AddItem "Received " & remote_dir & ExtractName(fName(II)) & " as " & local_dir & ExtractName(fName(II)) & " (" & numberof & " of " & UBound(fName) & ")  " & CurrDNFile & " bytes"
        WaitFor (226)
        numberof = numberof + 1
        DoEvents
    Next II
    
Size = 0
End Function
Function SetUpload()

Dim ret As Integer
trans = 0
halt_transfer = False
dir_info = ""
client.transferTotalBytes = 0
client.transferBytesSent = 0

  For i = LBound(fName) To UBound(fName)
    Open (fName(i)) For Binary As #10
    
    client.total_size = LOF(10)
    
    client.currentFile = fName(i)
    
    ProgressBar1.Max = SizeIt / 100
    Doing_Upload = True
    heard = True
    Call ipeer
    con.SendData "PWD" + Chr(13) + Chr(10)
    WaitFor (257)
    ServerLog.AddItem "Doing Upload"
    con.SendData prts + Chr(13) + Chr(10)
    ServerLog.AddItem prts
    WaitFor (200)
    con.SendData "STOR " & ExtractName(fName(i)) + Chr(13) + Chr(10)
    WaitFor (150)
    ChDir fPath(i)
  
    SendFile
    
    WaitFor (226)
  Call Wait(0.25)
  DoEvents
  Next

ProgressBar1.Position = 0
doing_multi = False
Size = 0

End Function
Function SendFile()
Dim temp As String

If doing_multi = False Then
cmdCancel.Enabled = True
End If

If halt_transfer = True Then
ProgressBar1.Position = 0
Exit Function
End If

Buffer = 4096

client.total_size = LOF(10)

With client

If Buffer > (.total_size - .transferBytesSent) Then
            Buffer = (.total_size - .transferBytesSent)
        End If
temp = Space$(Buffer)
        Get 10, , temp
.transferBytesSent = .transferBytesSent + Buffer
.transferTotalBytes = .transferTotalBytes + Buffer
trans = trans + .transferTotalBytes
trans = trans / 100
ProgressBar1.Position = trans

End With

data.SendData temp

End Function
Function Connect(IP As String)
   
   Call ipeer
   
With con
   .Close
   .RemoteHost = IP
   .RemotePort = ServPort
   .LocalPort = Int(Rnd * 99) + 5
   .Connect
End With

End Function

Sub ipeer()
On Error GoTo tryagain
data.Close
     LocalIP = con.LocalIP
      Do Until InStr(LocalIP, ".") = 0
           LocalIP = Left(LocalIP, InStr(LocalIP, ".") - 1) + "," + Right(LocalIP, Len(LocalIP) - InStr(LocalIP, "."))
       Loop
       

       Randomize Timer
       Nr1 = Int(Rnd * 12) + 5
       Nr2 = Int(Rnd * 254) + 1
       num1 = "PORT " + LocalIP + "," + Trim(Str(Nr1)) + "," + Trim(Str(Nr2))
       prts = num1
       data.LocalPort = (Nr1 * 256) + Nr2
       data.RemotePort = Trim(Str(Nr2)) ' was nr1
       data.Close
       data.Listen
       Exit Sub
       
tryagain:
Call Wait(0.25)
     data.Close
     LocalIP = con.LocalIP
      Do Until InStr(LocalIP, ".") = 0
           LocalIP = Left(LocalIP, InStr(LocalIP, ".") - 1) + "," + Right(LocalIP, Len(LocalIP) - InStr(LocalIP, "."))
       Loop
       

       Randomize Timer
       Nr1 = Int(Rnd * 12) + 5
       Nr2 = Int(Rnd * 254) + 1
       num1 = "PORT " + LocalIP + "," + Trim(Str(Nr1)) + "," + Trim(Str(Nr2))
       prts = num1
       data.LocalPort = (Nr1 * 256) + Nr2
       data.RemotePort = Trim(Str(Nr2)) ' was nr1
       data.Close
       data.Listen
End Sub
Sub Wait(WaitSeconds As Single)

Dim StartTime As Single

StartTime = Timer

Do While Timer < StartTime + WaitSeconds
DoEvents
Loop
End Sub
Sub WaitFor(WaitFor As String)
On Error Resume Next

Do While Response <> WaitFor
DoEvents
Loop

Response = 0
End Sub

Private Sub cmdCancel_Click()
If Doing_Upload = True Then
   con.SendData "ABOR " + Chr(13) + Chr(10)
   halt_transfer = True
Close #10
End If
If Doing_Download = True Then
   con.SendData "ABOR " + Chr(13) + Chr(10)
End If
End Sub
Private Sub Command2_Click()

If con.state = sckConnected Then
con.SendData "QUIT" + Chr(13) + Chr(10)
End If

Call Wait(1)

con.Close
data.Close

   End
End Sub

Private Sub con_DataArrival(ByVal bytesTotal As Long)
Dim tmps As String
Dim tri As Integer
Dim tre As String
Dim var As Variant
Dim leng As Integer
Dim ret As Boolean
Dim tmpArray() As String

    States(0).BackCode = "220"
    States(0).Command = "USER " + UserName
    States(1).BackCode = "331"
    States(1).Command = "PASS " + Password
    States(2).BackCode = "230"
    States(2).Command = "SYST"
    States(3).BackCode = "215"
    States(3).Command = ("CWD " + remote_dir + Chr(13) + Chr(10))
    
    
    
       con.GetData tmps, , bytesTotal
       
       leng = Len(tmps)
       leng = leng - 4
       tri = Len(tmps)
       tre = Mid(tmps, 1, (tri - 2))
       Response = Mid(tmps, 1, 3)
       
       If Mid(tmps, 1, 3) <> 220 Then  ' Server message and Help
       If Mid(tmps, 1, 3) = 214 Then
       GoTo skip
       End If
       If Mid(tmps, 1, 3) = 530 Then   ' Error message
       GoTo skip
       End If
       If Mid(tmps, 1, 3) = 257 Then   ' pwd message
       GoTo skip
       End If
       If Mid(tmps, 1, 3) = 221 Then   ' closing message
       GoTo skip
       End If
       ServerLog.AddItem tre
skip:
       Else
       If Left(tmps, 3) = "220" Then
       ' Server Welcome Message
           If InStr(1, Mid(tmps, 5, leng), vbCrLf) Then
           tre = vbCrLf & tre ' add one to beginning
           MessCnt = CountStr(tre, " ")
           
           If MessCnt = 0 Then Exit Sub
           
           ret = Parse2Array(tre, tmpArray(), vbCrLf)
           
           For t = 1 To MessCnt
           ServerLog.AddItem Mid(tmpArray(t), 2, Len(tmpArray(t)))
           If Mid(Mid(tmpArray(t), 2, Len(tmpArray(t))), 1, 10) = "220-Serv-U" Then ' SERV-U sends different path than War-ftp, mine is also compatable to war...,  go figure!
           servu = True
           End If
           Next
           Else
           ServerLog.AddItem tre
           End If
           End If
       End If
       
       If state < 4 Then
           If Left(tmps, 3) = States(state).BackCode Then
               bReplied = True
               con.SendData States(state).Command + Chr(13) + Chr(10)
               Debug.Print States(state).Command + Chr(13) + Chr(10)
               state = state + 1
               Exit Sub
           Else
           End If
           End If
           
           If Left(tmps, 4) = "LIST" Then
           Doing_list = True
           End If
           
           If Left(tmps, 4) = "150 " Then
            If Doing_Download = True Then
               Size = Val(Right(tmps, Len(tmps) - InStr(tmps, "(")))
               CurrDNFile = Size
               ProgressBar1.Max = Size
               open_file (fName(1))
               End If
    
               Exit Sub
skipper:
           End If
 
           If Left(tmps, 4) = "226 " Then
           Close #4
           data.Close
           heard = False
           Me.MousePointer = 0
           
           If doing_multi = False Then
           client.transferTotalBytes = 0
           trans = 0
           End If
           
           ProgressBar1.Position = 0
           ServerLog.AddItem "Done..."
           cmdCancel.Enabled = False
           
           If halt_transfer = True Then
           ServerLog.AddItem "226 Transfer aborted."
           halt_transfer = False
           Exit Sub
           End If
           End If
    
           If Left(tmps, 4) = "257 " Then
    
           If servu = False Then
           local_dir = Mid(tmps, 6, Len(tmps) - 30)
           ServerLog.AddItem "257 " & Mid(tmps, 6, Len(tmps) - 30) & " is working directory"
           End If
           
           If servu = True Then
           If Len(tmps) <= 30 Then
           local_dir = Mid(tmps, 6, 1)
           GoTo skiper
           End If
           
           local_dir = Mid(tmps, 7, Len(tmps) - 30)
           ServerLog.AddItem "257 " & local_dir & " is working directory"
           GoTo jump
skiper:
           ServerLog.AddItem "257 " & Mid(tmps, 6, 1) & " is working directory"
           End If
           
jump:
           End If
           
           
           If Left(tmps, 4) = "530 " Then
           data.Close
           ServerLog.AddItem "530 password incorrect, not logged in"
           Exit Sub
           End If
           
           If Left(tmps, 4) = "425 " Then
           data.Close
           con.SendData "TYPE A" + Chr(13) + Chr(10)
           WaitFor (200)
           con.SendData ("PWD " + Chr(13) + Chr(10))
           End If
           
           If Left(tmps, 4) = "426 " Then
           data.Close
           End If
           
           
           If Left(tmps, 4) = "550 " Then
           Me.MousePointer = 0
           
           If Doing_Download = True Then
           Doing_Download = False
           Close #1
           data.Close
           Exit Sub
           End If
           
           If Doing_Upload = True Then
           Close #10
           Doing_Upload = False
           data.Close
           Exit Sub
           End If
           
           If Doing_Download = True Then
           Doing_Download = False
           Close #1
           data.Close
           End If
           
           data.Close
           Me.MousePointer = 0
           End If
End Sub

Private Sub data_Close()
    data.Close
    Close #1
    Me.MousePointer = 0
    ProgressBar1.Position = 0
    ServerLog.AddItem "Data connection closed"
End Sub

Private Sub data_ConnectionRequest(ByVal requestID As Long)
   data.Close
   data.Accept requestID
End Sub

Private Sub data_DataArrival(ByVal bytesTotal As Long)
Dim rmdata As String
cmdCancel.Enabled = True
On Error Resume Next
data.GetData rmdata, , bytesTotal
trans = trans + bytesTotal
client.transferTotalBytes = trans
ProgressBar1.Position = trans
'
Put #1, , rmdata
End Sub

Private Sub data_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    data.Close
End Sub

Private Sub data_SendComplete()
   With client
                If .total_size = .transferBytesSent Then
                    Close #10
                    Close #4
                    data.Close
                    .transferBytesSent = 0
                    ProgressBar1.Position = 0
                    cmdCancel.Enabled = False
                Else
                    SendFile
                End If
    End With
End Sub

Private Sub Form_Load()
  Set inigo = New cIniFile
  Retrned = inigo.LastReturnCode
  inigo.Path = App.Path & "\autoup.ini"
 Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set inigo = Nothing
End Sub

Private Sub Timer1_Timer()
Dim ret As Boolean
Timer1.Enabled = False

' true for upload
 ret = DoTheIni(True)
 
 If ret = False Then GoTo air
 
 UserName = "user_account_name"
 Password = "password_for_account"
 
 ' usually 21 for ftp server
 ' but could be any
 ServPort = 21
 
 Connect "127.0.0.1"
 
 WaitFor "250"
 
 SetUpload
 'SetDownload
 
 If Check1.Value = 1 Then
 Timer2.Enabled = True
 End If
 
 Exit Sub
 
air:
ServerLog.Clear
ServerLog.AddItem "No files found to transfer"
End Sub

Private Sub Timer2_Timer()
 If heard = False And Check1.Value = 1 Then
 Call Command2_Click
 End If
End Sub
