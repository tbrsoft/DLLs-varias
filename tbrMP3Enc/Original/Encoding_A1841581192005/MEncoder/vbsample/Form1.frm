VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 Encoder"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      ItemData        =   "Form1.frx":0E42
      Left            =   60
      List            =   "Form1.frx":0E6D
      TabIndex        =   10
      Text            =   "BitRate"
      Top             =   600
      Width           =   2535
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Private"
      Height          =   195
      Left            =   2850
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "CRC"
      Height          =   195
      Left            =   2100
      TabIndex        =   8
      Top             =   1080
      Width           =   645
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Original"
      Height          =   195
      Left            =   1140
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1200
      TabIndex        =   6
      Top             =   1950
      Width           =   915
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Copyright"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar prog 
      Height          =   345
      Left            =   60
      TabIndex        =   3
      Top             =   1500
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Encode"
      Height          =   345
      Left            =   90
      TabIndex        =   2
      Top             =   1950
      Width           =   915
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   150
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   450
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   2850
      TabIndex        =   0
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   3180
      TabIndex        =   11
      Top             =   1980
      Width           =   465
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   3195
      TabIndex        =   4
      Top             =   660
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'API declarations for encoding wrapper
Private Declare Function SetBitrate Lib "MEncoder.dll" (ByVal bit As Integer) As Long
Private Declare Function EncodeMp3 Lib "MEncoder.dll" (ByVal lpszWavFile As String, lpCallback As Any) As Long
Private Declare Function SetCopyright Lib "MEncoder.dll" (ByVal cpy As Boolean) As Long
Private Declare Function SetOriginal Lib "MEncoder.dll" (ByVal org As Boolean) As Long
Private Declare Function SetCRC Lib "MEncoder.dll" (ByVal crc As Boolean) As Long
Private Declare Function SetPrivate Lib "MEncoder.dll" (ByVal priv As Boolean) As Long
Private Declare Function Cancel Lib "MEncoder.dll" (ByVal cncl As Boolean) As Long

'/////////////////////////////////////////////////////////////////////////////
'  Acceptable Bitrates                                                      //
' 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256 and 320 allowed  //
'/////////////////////////////////////////////////////////////////////////////

Private Sub Combo1_Click()

    'if the first item in teh drop down box was selected then
    'set the bitrate (kbps). this one will be 320
    If Combo1.ListIndex = 0 Then
        Call SetBitrate(320)
        'end this routine "Its a optimization thing" ;)
        Exit Sub
    End If

    '2nd item
    If Combo1.ListIndex = 1 Then
        Call SetBitrate(256)
        Exit Sub
    End If
    
    '3rd item
    If Combo1.ListIndex = 2 Then
        Call SetBitrate(224)
        Exit Sub
    End If
    
    '4th item
    If Combo1.ListIndex = 3 Then
        Call SetBitrate(192)
        Exit Sub
    End If
    
    '5th item
    If Combo1.ListIndex = 4 Then
        Call SetBitrate(160)
        Exit Sub
    End If
    
    '6th item
    If Combo1.ListIndex = 5 Then
        Call SetBitrate(128)
        Exit Sub
    End If
    
    '7th item
    If Combo1.ListIndex = 6 Then
        Call SetBitrate(112)
        Exit Sub
    End If
    
    '8th item
    If Combo1.ListIndex = 7 Then
        Call SetBitrate(96)
        Exit Sub
    End If
    
    '9th item
    If Combo1.ListIndex = 8 Then
        Call SetBitrate(80)
        Exit Sub
    End If
    
    '10th item
    If Combo1.ListIndex = 9 Then
        Call SetBitrate(64)
        Exit Sub
    End If
    
    '11th item
    If Combo1.ListIndex = 10 Then
        Call SetBitrate(56)
        Exit Sub
    End If
    
        '12th item
    If Combo1.ListIndex = 11 Then
        Call SetBitrate(48)
        Exit Sub
    End If
    
        '13th item
    If Combo1.ListIndex = 12 Then
        Call SetBitrate(40)
        Exit Sub
    End If
    
End Sub

Private Sub Command1_Click()

    ' set teh commondialog to filter only wave files
    CommonDialog1.Filter = "Wave Audio|*.wav"
    'open teh dialog
    CommonDialog1.ShowOpen
    'set teh text box to the filename of the wave fiel picked
    Text1.Text = CommonDialog1.FileName
    
End Sub

Private Sub Command2_Click()

    Dim nRes As Integer
    
    'If No File is selected then don't encode or crashola :)
    If Text1.Text = vbNullString Then Exit Sub
    
    'set file info extras :)
    If Check1.Value = 0 Then
        Call SetCopyright(False)
    End If
    
    If Check1.Value = 1 Then
        Call SetCopyright(True)
    End If
    
    
    If Check2.Value = 0 Then
        Call SetOriginal(False)
    End If
    
    If Check2.Value = 1 Then
        Call SetOriginal(True)
    End If
    
    
    If Check3.Value = 0 Then
        Call SetCRC(False)
    End If
    
    If Check3.Value = 1 Then
        Call SetCRC(True)
    End If
    
    
    If Check4.Value = 0 Then
        Call SetPrivate(False)
    End If
    
    If Check4.Value = 1 Then
        Call SetPrivate(True)
    End If


    'Now start emcpding the file
    nRes = EncodeMp3(Text1.Text, AddressOf EnumEncoding)
    
    'User Notifications
    If nRes <> -1 Then
        MsgBox "MP3 encoding complete", vbInformation, App.Title
    ElseIf nRes = -2 Then
        MsgBox "Encoding stopped by user", vbExclamation, App.Title
    Else
        MsgBox "Encoding failed", vbExclamation, App.Title
    End If
    
    'we are done so reset % labels and prog bar :)
    
    lblPercent.Caption = "0%"
    
    Label1.Caption = "0%"
    
    prog.Value = 0
    
End Sub


Private Sub Command3_Click()

    'cancel is called
    Call Cancel(True)
    
End Sub

Private Sub Form_Load()

    'set defaults
    
    Call SetBitrate(192)
    Call SetCopyright(False)
    Call SetOriginal(False)
    Call SetCRC(False)
    Call SetPrivate(False)
    
End Sub
