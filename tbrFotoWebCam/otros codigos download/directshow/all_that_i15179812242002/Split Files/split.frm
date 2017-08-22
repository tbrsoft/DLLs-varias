VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "File Spliter Pro"
   ClientHeight    =   2028
   ClientLeft      =   48
   ClientTop       =   624
   ClientWidth     =   6828
   Icon            =   "split.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2028
   ScaleWidth      =   6828
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   132
      Left            =   960
      TabIndex        =   12
      Top             =   1750
      Width           =   4932
      _ExtentX        =   8700
      _ExtentY        =   233
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   132
      Left            =   960
      TabIndex        =   11
      Top             =   1600
      Width           =   4932
      _ExtentX        =   8700
      _ExtentY        =   233
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6000
      TabIndex        =   10
      Top             =   1560
      Width           =   732
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SPLIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   732
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   564
      Left            =   120
      TabIndex        =   8
      Top             =   420
      Width           =   6612
      _ExtentX        =   11663
      _ExtentY        =   995
      _Version        =   393216
      LargeChange     =   10
      Min             =   2
      Max             =   100
      SelStart        =   2
      TickStyle       =   2
      TickFrequency   =   2
      Value           =   2
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   0
      Top             =   240
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Filter          =   "All Files|*.*"
      Flags           =   4101
   End
   Begin VB.Label Label3 
      Caption         =   "Number of Files :"
      Height          =   252
      Left            =   5010
      TabIndex        =   7
      Top             =   1170
      Width           =   1212
   End
   Begin VB.Label text5 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   6240
      TabIndex        =   6
      Top             =   1140
      Width           =   492
   End
   Begin VB.Label text4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   3690
      TabIndex        =   5
      Top             =   1140
      Width           =   1092
   End
   Begin VB.Label text3 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   2970
      TabIndex        =   4
      Top             =   1140
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "Split Size :"
      Height          =   252
      Left            =   2160
      TabIndex        =   3
      Top             =   1170
      Width           =   732
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   6612
   End
   Begin VB.Label text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   840
      TabIndex        =   1
      Top             =   1140
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "File Size :"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   1170
      Width           =   732
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      Begin VB.Menu mopen 
         Caption         =   "&Open"
      End
      Begin VB.Menu msplit 
         Caption         =   "&Split"
      End
      Begin VB.Menu mexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim file$

Private Sub Command1_Click()
  On Error GoTo errorhandler
  Const maxalloc = 1048576
  Dim curfile As Long, filesize As Long, numfile As Long, z As Long, num$, file2$
  Slider1.Enabled = False
  Command1.Enabled = False
  msplit.Enabled = False
  mopen.Enabled = False
  ProgressBar2.Value = 0
  filesize = text4.Caption
  numfile = text5.Caption
  Open file$ For Binary Access Read Lock Write As #1 Len = 1
  For curfile = 1 To numfile
    num$ = Mid$(Str$(curfile - 1), 2)
    If (curfile - 1) < 10 Then num$ = "0" & num$
    file2$ = file$ & ".S" & num$
    Kill file2$
    Open file2$ For Binary Access Read Write Lock Write As #2 Len = 1
    ProgressBar1.Value = 0
    z = filesize
    Do
      If z < maxalloc Then
        ReDim s(z) As Byte
      Else
        ReDim s(maxalloc) As Byte
      End If
      If (curfile = numfile) And (LOF(1) - Loc(1)) < z And (LOF(1) - Loc(1)) < maxalloc Then ReDim s(LOF(1) - Loc(1)) As Byte
      Get #1, , s
      DoEvents
      Put #2, , s
      z = z - maxalloc
      ProgressBar1.Value = (Loc(2) / filesize) * 100
    Loop Until z <= 0
    ProgressBar1.Value = 100
    ProgressBar2.Value = (curfile / numfile) * 100
    Close #2
  Next curfile
  Close #1
  Slider1.Enabled = True
  Command1.Enabled = True
  msplit.Enabled = True
  mopen.Enabled = True
  Exit Sub
errorhandler:
  If Err.Number = 53 Then
    Resume Next
  Else
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Err.Description, 16, "Error", Err.HelpFile, Err.HelpContext
  End If
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Form_Load()
  cmd1.FileName = ""
  file$ = ""
  cmd1.ShowOpen
  file$ = cmd1.FileName
  If file$ = "" Then End
  text1.Caption = file$
  text2.Caption = FileLen(file$)
  If text2.Caption < 2 Then
    MsgBox "File Too Small !", vbExclamation, "ERROR"
    End
  End If
  calcul
End Sub

Private Sub mexit_Click()
  End
End Sub

Private Sub mopen_Click()
  cmd1.FileName = ""
  cmd1.ShowOpen
  If cmd1.FileName = "" Then Exit Sub
  If FileLen(cmd1.FileName) < 2 Then
    MsgBox "File Too Small !", vbExclamation, "ERROR"
    Exit Sub
  End If
  file$ = cmd1.FileName
  text1.Caption = file$
  text2.Caption = FileLen(file$)
  calcul
End Sub

Private Sub msplit_Click()
  Command1_Click
End Sub

Private Sub Slider1_Change()
  calcul
End Sub

Private Sub Slider1_Scroll()
  calcul
End Sub

Private Sub calcul()
  text4.Caption = calculmod(text2.Caption, Slider1.Value)
  text5.Caption = calculmod(text2.Caption, text4.Caption)
  text3.Caption = Left$(((text4.Caption / text2.Caption) * 100), 4) & " %"
End Sub

Function calculmod(var1 As Long, var2 As Long) As Long
  Dim vartemp As Long
  vartemp = CLng(var1 / var2)
  If CLng(var1 / var2) > (var1 / var2) Then vartemp = vartemp - 1
  If (var1 Mod var2) <> 0 Then vartemp = vartemp + 1
  calculmod = vartemp
End Function
