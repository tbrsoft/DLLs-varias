VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super"
   ClientHeight    =   7116
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8232
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7116
   ScaleWidth      =   8232
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   600
      Width           =   5172
   End
   Begin VB.TextBox Text8 
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1320
      Width           =   5172
   End
   Begin VB.TextBox Text7 
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2040
      Width           =   5172
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Get Memory Status"
      Height          =   492
      Left            =   360
      TabIndex        =   19
      Top             =   3720
      Width           =   1932
   End
   Begin VB.CommandButton Command4 
      Caption         =   "is Key Down ?"
      Height          =   492
      Left            =   360
      TabIndex        =   18
      Top             =   2880
      Width           =   1932
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Browse For Folder"
      Height          =   492
      Left            =   360
      TabIndex        =   17
      Top             =   2040
      Width           =   1932
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get All Drives Type && Free Space"
      Height          =   492
      Left            =   360
      TabIndex        =   16
      Top             =   4560
      Width           =   1932
   End
   Begin VB.TextBox Text6 
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4200
      Width           =   5172
   End
   Begin VB.TextBox Text5 
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3480
      Width           =   5172
   End
   Begin VB.TextBox Text4 
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2760
      Width           =   5172
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4920
      Width           =   5172
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5640
      Width           =   5172
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6360
      Width           =   5172
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Flashing"
      Enabled         =   0   'False
      Height          =   492
      Left            =   360
      TabIndex        =   3
      Top             =   6240
      Width           =   1932
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1250
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Windows About"
      Height          =   492
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1932
   End
   Begin VB.CommandButton Command12 
      Caption         =   "About"
      Height          =   492
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1932
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Begin Flashing"
      Height          =   492
      Left            =   360
      TabIndex        =   0
      Top             =   5640
      Width           =   1932
   End
   Begin VB.Label Label9 
      Caption         =   "User \ Start Menu \ Programs"
      Height          =   252
      Left            =   2640
      TabIndex        =   25
      Top             =   360
      Width           =   5172
   End
   Begin VB.Label Label8 
      Caption         =   "User Documents"
      Height          =   252
      Left            =   2640
      TabIndex        =   23
      Top             =   1080
      Width           =   5172
   End
   Begin VB.Label Label7 
      Caption         =   "User Desktop Directory"
      Height          =   252
      Left            =   2640
      TabIndex        =   21
      Top             =   1800
      Width           =   5172
   End
   Begin VB.Label Label6 
      Caption         =   "Temp Directory"
      Height          =   252
      Left            =   2640
      TabIndex        =   15
      Top             =   3960
      Width           =   5172
   End
   Begin VB.Label Label5 
      Caption         =   "System Directory"
      Height          =   252
      Left            =   2640
      TabIndex        =   13
      Top             =   3240
      Width           =   5172
   End
   Begin VB.Label Label4 
      Caption         =   "Windows Directory"
      Height          =   252
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   5172
   End
   Begin VB.Label Label3 
      Caption         =   "User Name"
      Height          =   252
      Left            =   2640
      TabIndex        =   9
      Top             =   4680
      Width           =   5172
   End
   Begin VB.Label Label2 
      Caption         =   "Computer Name"
      Height          =   252
      Left            =   2640
      TabIndex        =   8
      Top             =   5400
      Width           =   5172
   End
   Begin VB.Label Label1 
      Caption         =   "Windows Version"
      Height          =   252
      Left            =   2640
      TabIndex        =   7
      Top             =   6120
      Width           =   5172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Timer1.Enabled = False
  Command1.Enabled = False
  Command11.Enabled = True
End Sub

Private Sub Command11_Click()
  Timer1.Enabled = True
  Command1.Enabled = True
  Command11.Enabled = False
End Sub

Private Sub Command12_Click()
  frmAbout.Show vbModal, Me
End Sub

Private Sub Command13_Click()
  ShowAbout App, Me
End Sub

Private Sub Command2_Click()
Dim qwe As String, q As String, t As Long, w As Long, f As Currency
For t = 0 To 25
  q = Chr$(Asc("A") + t) & ":"
  w = DriveType(q)
  If (w >= 2 And w <= 6) Or w = 0 Then
    qwe = qwe & vbCrLf & "Drive " & q & " IS : " & DriveTypeS(q)
'   If w = 3 Then
      f = FreeSpace(q)
      If f >= 0 Then
        qwe = qwe & "   Free Space : " & Calc(f)
      ElseIf w = 2 Or w = 5 Then
        qwe = qwe & "   NO DISC"
      End If
'   End If
  End If
Next t
MsgBox Right$(qwe, Len(qwe) - 2), vbInformation, "Drives Type & Free Space"
End Sub

Private Sub Command3_Click()
  MsgBox BrowseForFolder("Select A Directory", "", Me), vbInformation, "Folder Selected"
End Sub

Private Sub Command4_Click()
  frmKeyDown.Show vbModal, Me
End Sub

Private Sub Command5_Click()
  Dim qwe As MemoryStatus, q As String
  qwe = GetMemory
  q = "Total Memory : " & Calc(qwe.TotalPhys) & vbCrLf _
    & "Free Memory : " & Calc(qwe.AvailPhys) & vbCrLf _
    & "Memory Load : " & Round(qwe.MemoryLoad2, 2) & "%"
  MsgBox q, vbInformation, "Memory Status"
End Sub

Private Sub Form_Load()
  Me.Icon = LoadResPicture(101, vbResIcon)
  Text1.Text = GetWindowsVersion.dwFullTextV
  Text2.Text = GetComputerName
  Text3.Text = GetUserName
  Text4.Text = GetWindowsDir
  Text5.Text = GetSystemDir
  Text6.Text = GetTempDir
  Text7.Text = GetSpecialFolder(DIR_USER_DESKTOP)
  Text8.Text = GetSpecialFolder(DIR_USER_MY_DOCUMENTS)
  Text9.Text = GetSpecialFolder(DIR_USER_START_MENU_PROGRAMS)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub Timer1_Timer()
  Flash Me
End Sub

Private Function Calc(ByVal var1 As Currency) As String
  Dim Z As Byte, q As String
  q = ""
  Z = 0
  While var1 >= 1000 And Z < 3
    var1 = var1 / 1024
    Z = Z + 1
  Wend
  q = q & Round(var1, 2)
  If Z = 3 Then
    q = q & " G."
  ElseIf Z = 2 Then
    q = q & " MB"
  ElseIf Z = 1 Then
    q = q & " KB"
  ElseIf Z = 0 Then
    q = q & " Bytes"
  End If
  Calc = q
End Function
