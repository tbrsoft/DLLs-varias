VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transparent"
   ClientHeight    =   4332
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4692
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4332
   ScaleWidth      =   4692
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Fade To Trans Value"
      Height          =   492
      Left            =   2400
      TabIndex        =   9
      Top             =   3000
      Width           =   1932
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set Transparent Value"
      Height          =   492
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   1932
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Make &Transparent"
      Height          =   492
      Left            =   2400
      TabIndex        =   7
      Top             =   2400
      Width           =   1932
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Make &Opaque"
      Height          =   492
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1932
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Get Transparent Value"
      Height          =   492
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1932
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Transparent ?"
      Height          =   492
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1932
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TRANSPARENT FORM"
      Height          =   492
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1932
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FORM  &&  &BUTTONS FROM PICTURES"
      Height          =   492
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1932
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SHAPED FORM"
      Height          =   492
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FORM  WITH  &MASK"
      Height          =   492
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1932
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   252
      Left            =   3000
      TabIndex        =   11
      Top             =   3720
      Width           =   252
      _ExtentX        =   445
      _ExtentY        =   445
      _Version        =   393216
      Value           =   255
      BuddyControl    =   "Label1"
      BuddyDispid     =   196619
      OrigLeft        =   1800
      OrigTop         =   2760
      OrigRight       =   2052
      OrigBottom      =   3012
      Increment       =   16
      Max             =   255
      Min             =   31
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65537
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      Height          =   252
      Left            =   2640
      TabIndex        =   10
      Top             =   3720
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transparent Value :"
      Height          =   192
      Left            =   1080
      TabIndex        =   12
      Top             =   3720
      Width           =   1404
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Command1_Click()
  frmMask.Show
End Sub

Private Sub Command10_Click()
  MsgBox isTransparent(Me)
End Sub

Private Sub Command2_Click()
  frmButtons.Show
End Sub

Private Sub Command3_Click()
  frmShapes.Show
End Sub

Private Sub Command4_Click()
  frmTrans.Show
End Sub

Private Sub Command5_Click()
  Dim qwe As String
  If isTransparent(Me) = LWA_COLORKEY Then
    qwe = Hex(GetTrans(Me))
    While Len(qwe) < 6
      qwe = "0" & qwe
    Wend
    MsgBox "BGR : " & qwe
  Else
    MsgBox GetTrans(Me)
  End If
End Sub

Private Sub Command6_Click()
  Me.BackColor = &H8000000F
  SetTrans Me, UpDown1.Value
End Sub

Private Sub Command7_Click()
  MakeOpaque Me
  Me.BackColor = &H8000000F
End Sub

Private Sub Command8_Click()
  Me.BackColor = &H8000000F
  FadeTo Me, UpDown1.Value
End Sub

Private Sub Command9_Click()
  Me.BackColor = &HFF00FF
  MakeTrans Me
End Sub

Private Sub Form_Load()
  FadeIn Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.BackColor = &H8000000F
  FadeOut Me
  End
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub
