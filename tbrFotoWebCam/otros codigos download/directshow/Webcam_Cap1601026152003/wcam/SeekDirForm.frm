VERSION 5.00
Begin VB.Form frmSeekDir 
   Caption         =   "Images Directory"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SeekDirForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1530
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   3330
      Width           =   1365
   End
   Begin VB.CommandButton ActionButton 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   135
      TabIndex        =   2
      ToolTipText     =   "Set images directory"
      Top             =   3330
      Width           =   1320
   End
   Begin VB.DriveListBox Drive1 
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2760
   End
   Begin VB.DirListBox Dir1 
      Height          =   2640
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2760
   End
End
Attribute VB_Name = "frmSeekDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActionButton_Click()
'
bIDOKPressed = True
frmSettings.tFilePath.Text = Me.Dir1.Path
Me.Hide
Unload Me
End Sub

Private Sub Cancel_Click()
Call Form_Terminate
End Sub

Private Sub Form_Load()
'
On Error Resume Next
'
Dim sCurrPath As String
'
sCurrPath = frmSettings.tFilePath.Text
bIDOKPressed = False
If Dir(sCurrPath, vbDirectory) <> "" Then
 Me.Dir1.Path = sCurrPath
Else
 Me.Dir1.Path = App.Path
End If
Me.SetFocus
End Sub

Private Sub Form_Resize()
'
Dim lMinWidth As Long
Dim lMinHeight As Long
Dim lFWidth As Long
Dim lFHeight As Long
'
If Me.WindowState = 0 Then
 lMinWidth = 3120
 lMinHeight = 4365
 If Me.Width < lMinWidth Then
  Me.Width = lMinWidth
 End If
 If Me.Height < lMinHeight Then
  Me.Height = lMinHeight
 End If
 '
 lFWidth = Me.Width
 lFHeight = Me.Height
 '
 Me.Drive1.Width = lFWidth - 360
 '
 Me.Dir1.Width = lFWidth - 360
 Me.Dir1.Height = lFHeight - 1725
 '
 Me.ActionButton.Top = lFHeight - 1035
 '
 Me.Cancel.Top = lFHeight - 1035
 Me.Cancel.Left = lFWidth - 1590
End If
End Sub

Private Sub Form_Terminate()
Me.Hide
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Form_Terminate
End Sub
