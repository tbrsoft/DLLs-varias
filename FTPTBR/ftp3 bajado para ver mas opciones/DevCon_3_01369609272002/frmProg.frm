VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data transfer"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2760
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2880
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Download"
      Height          =   375
      Left            =   1493
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2813
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1485
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Time remaining:"
      Height          =   195
      Left            =   3480
      TabIndex        =   20
      Top             =   240
      Width           =   1410
   End
   Begin VB.Label lbTimeLeft 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4560
      TabIndex        =   19
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lbSpeed 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4200
      TabIndex        =   18
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lbTime 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4560
      TabIndex        =   17
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Speed:"
      Height          =   195
      Left            =   3480
      TabIndex        =   16
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   195
      Left            =   3480
      TabIndex        =   15
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lbCount 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1160
      TabIndex        =   10
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label lbCelkem 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lbLeft 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbFar 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Bytes to send:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Bytes sent:"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kb to transfer:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File:             /5"
      Height          =   255
      Left            =   375
      TabIndex        =   3
      Top             =   1260
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Downloading file:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   735
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   120
      Picture         =   "frmProg.frx":6246
      Top             =   1220
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frmProg.frx":62C8
      Top             =   720
      Width           =   240
   End
End
Attribute VB_Name = "frmProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const sReadBuffer = 1024
Private hFile As Long

Private Sub Command1_Click()
If hFile <> 0 Then
InternetCloseHandle hFile
MsgBox "Opertion Aborted.", vbInformation
End If
End Sub

Private Sub Download()
   Dim sBuffer As String
   Dim FileData As String
   Dim Ret As Long, SentBytes As Long, sAllBytes As Long, z As Long
   Dim i As Integer, FF As Integer
   Dim Kam As String, Ode As String
   Dim Fs As Long, StartT As Long, t As Long, Cnt As Long, p As Long
   Dim spRate As Single

z = lbCelkem.Caption
spRate = 0
sAllBytes = 0
p = 0
With ProgressBar2
   .Value = 0
   .Min = 0
   .Max = z * 1.024
End With
lbLeft.Visible = True
lbLeft.Caption = z / 1024
lbLeft.Refresh
lbCelkem.Caption = Format(z, "### ### ###.##") & " Kb"
Command1.Enabled = True
    For i = 0 To List1.ListCount - 1
           Fs = List2.List(i)
           With ProgressBar1
                .Value = 0
                .Min = 0
                .Max = Fs * 1.024
            End With
            Label1.Caption = "Downloading File: " & List1.List(i) & " / " & Fs & " bytes."
            Label1.Refresh
            lbCount.Caption = i + 1
            lbCount.Refresh
            Ode = Klic & List1.List(i)
            Kam = strPath & List1.List(i)
            frmmain.txtInfo.SelText = Time & " >Finishing File Transfer..." & vbCrLf & " > Downloading from: " & Ode & ", to: " & Kam & vbCrLf
            frmmain.StatusBar1.Panels(2).Text = "Transfering File.Please Wait....."
              hFile = FtpOpenFile(server, Ode, GENERIC_READ, FTP_TRANSFER_TYPE_BINARY, 0)
                 If hFile = 0 Then
                   MsgBox "Can't open file path!", vbExclamation, "Invaild URL"
                   frmmain.txtInfo.SelText = Time & " > Can't open file path! Request canceled!" & vbCrLf
                  Exit Sub
            End If
              sBuffer = Space(sReadBuffer)
              FileData = ""
              SentBytes = 0
              StartT = GetTickCount

                Do
                    InternetReadFile hFile, sBuffer, sReadBuffer, Ret
                    If Ret <> sReadBuffer Then
                        sBuffer = Left$(sBuffer, Ret)
                    End If
                    FileData = FileData + sBuffer
                    SentBytes = SentBytes + Ret
                    sAllBytes = sAllBytes + Ret
                    lbFar.Caption = Format(sAllBytes / 1024, "### ### ###.##") & " Kb"
                    lbFar.Refresh
                    lbLeft.Caption = Format((z / 1000) - (sAllBytes / 1024), "### ### ###.##") & " Kb"
                    lbLeft.Refresh
                        If SentBytes <> 0 Then
                            t = GetTickCount - StartT
                            If t <> 0 Then
                                spRate = (spRate + ((SentBytes / 1000) / (t / 1000))) / 2
                                lbTime.Caption = Format(p + (t / 1000), "### ###") & " s" 'CStr(Int(((Fs - SentBytes) / 1000) / spRate)) & " s"
                                lbTime.Refresh
                                lbSpeed.Caption = Format(spRate, "#.##") & " Kbps"
                                lbSpeed.Refresh
                                lbTimeLeft.Caption = Format(((z / 1000) - (sAllBytes / 1024)) / spRate, "### ###") & " s"
                                lbTimeLeft.Refresh
                            End If
                        End If
                    ProgressBar1.Value = SentBytes
                    ProgressBar2.Value = sAllBytes
    Loop Until Ret <> sReadBuffer
             FF = FreeFile
             Open Kam For Binary As #FF
                 Put #FF, , FileData
             Close #FF
    p = t / 1000
    InternetCloseHandle hFile
    frmmain.txtInfo.SelText = Time & " > OK" & vbCrLf
    Next i
MsgBox "File Transfer Completed.", vbInformation
Unload Me
End Sub

Private Sub Upload()
    Dim Cnt As Long, nFileLen As Long, nRet As Long, nTotFileLen As Long
    Dim sBuffer As String * 1024
   Dim Ret As Long, SentBytes As Long, sAllBytes As Long, z As Long
   Dim i As Integer
   Dim Kam As String, Ode As String
   Dim Fs As Long, StartT As Long, t As Long, p As Long
   Dim spRate As Single

z = lbCelkem.Caption
spRate = 0
sAllBytes = 0
p = 0
With ProgressBar2
   .Value = 0
   .Min = 0
   .Max = z
End With
lbLeft.Visible = True
lbLeft.Caption = z / 1024
lbLeft.Refresh
lbCelkem.Caption = Format(z, "### ### ###.##") & " Kb"
Command1.Enabled = True
    For i = 0 To List1.ListCount - 1
           Fs = List2.List(i)
           With ProgressBar1
                .Value = 0
                .Min = 0
                .Max = Fs
            End With
            Label1.Caption = "Uploading File: " & List1.List(i) & " / " & Fs & " bytes."
            Label1.Refresh
            lbCount.Caption = i + 1
            lbCount.Refresh
            Ode = strPath & List1.List(i)
            Kam = Klic & List1.List(i)
            frmmain.txtInfo.SelText = Time & " > Begining File Transfer..." & vbCrLf & " > Uploading from: " & Ode & ", to: " & Kam & vbCrLf
            frmmain.StatusBar1.Panels(2).Text = "Uploading Data. Please Wait...."
            hFile = FtpOpenFile(server, Kam, GENERIC_WRITE, FTP_TRANSFER_TYPE_BINARY, 0)
            If hFile = 0 Then
                MsgBox "Cant create requested file on server", vbExclamation, App.Title
                frmmain.txtInfo.SelText = Time & " > Cant create requested file on server! request canceled!" & vbCrLf
                Exit Sub
            End If
            SentBytes = 0
            nFileLen = 0
            StartT = GetTickCount
            Open Ode For Binary As #1
                nTotFileLen = LOF(1)
                Do
                    Get #1, , sBuffer
                    If nFileLen < nTotFileLen - sReadBuffer Then
                        If InternetWriteFile(hFile, sBuffer, sReadBuffer, nRet) = 0 Then
                            MsgBox "Could Not Write File", vbExclamation, App.Title
                            frmmain.txtInfo.SelText = Time & " > Could Not Write File!" & vbCrLf
                            Exit Do
                        End If
                        SentBytes = SentBytes + sReadBuffer
                        sAllBytes = sAllBytes + sReadBuffer
                        nFileLen = nFileLen + sReadBuffer
                    Else
                        If InternetWriteFile(hFile, sBuffer, nTotFileLen - nFileLen, nRet) = 0 Then
                            MsgBox "Could Not Write File!", vbExclamation, App.Title
                            frmmain.txtInfo.SelText = Time & " >Could Not Write File!" & vbCrLf
                            Exit Do
                        End If
                        SentBytes = SentBytes + (nTotFileLen - nFileLen)
                        sAllBytes = sAllBytes + (nTotFileLen - nFileLen)
                        nFileLen = nTotFileLen
                    End If
                    lbFar.Caption = Format(sAllBytes / 1024, "### ### ###.##") & " Kb"
                    lbFar.Refresh
                    lbLeft.Caption = Format((z / 1000) - (sAllBytes / 1024), "### ### ###.##") & " Kb"
                    lbLeft.Refresh
                        If SentBytes <> 0 Then
                            t = GetTickCount - StartT
                            If t <> 0 Then
                                spRate = (spRate + ((SentBytes / 1000) / (t / 1000))) / 2
                                lbTime.Caption = Format(p + (t / 1000), "### ###") & " s" 'CStr(Int(((Fs - SentBytes) / 1000) / spRate)) & " s"
                                lbTime.Refresh
                                lbSpeed.Caption = Format(spRate, "#.##") & " Kbps"
                                lbSpeed.Refresh
                                lbTimeLeft.Caption = Format(((z / 1000) - (sAllBytes / 1024)) / spRate, "### ###") & " s"
                                lbTimeLeft.Refresh
                            End If
                        End If
                    ProgressBar1.Value = nFileLen
                    ProgressBar2.Value = sAllBytes
                Loop Until nFileLen >= nTotFileLen
            Close
            p = t / 1000
            InternetCloseHandle hFile
            frmmain.txtInfo.SelText = Time & " > OK" & vbCrLf
    Next i
MsgBox "Data transfer completed.", vbInformation
Unload Me
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Upload" Then
    Upload
Else: Download
End If
End Sub

Private Sub Form_UnLoad(Cancel As Integer)
List1.Clear
End Sub

Private Sub Form_Load()
With ProgressBar1
         .Value = 0
         .Min = 0
End With
With ProgressBar2
         .Value = 0
         .Min = 0
End With
End Sub

