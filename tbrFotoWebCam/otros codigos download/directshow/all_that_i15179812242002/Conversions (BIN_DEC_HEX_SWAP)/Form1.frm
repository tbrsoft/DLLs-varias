VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conversions & Swap"
   ClientHeight    =   4332
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4332
   ScaleWidth      =   3000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3840
      Width           =   372
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3480
      Width           =   372
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Swap var1 && var2"
      Height          =   648
      Left            =   1560
      TabIndex        =   8
      Top             =   3480
      Width           =   1212
   End
   Begin VB.CommandButton Command8 
      Caption         =   "High Word && Low Word"
      Height          =   732
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   1212
   End
   Begin VB.CommandButton Command7 
      Caption         =   "High Byte && Low Byte"
      Height          =   732
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1212
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Bin To Hex"
      Height          =   612
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hex To Bin"
      Height          =   612
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dec To Bin"
      Height          =   612
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1212
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Bin To Dec"
      Height          =   612
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hex To Dec"
      Height          =   612
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dec To Hex"
      Height          =   612
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "var2 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   612
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "var1 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim var1 As Single, var2 As Single

Private Sub Command1_Click()
  On Local Error GoTo ErrHnd
  Dim qwe As Long, qaz As String
Debut:
  qaz = InputBox("Enter a Long : -2147483648 to 2147483647", "Dec To Hex", 0)
  If qaz <> "" Then
    qwe = qaz
    MsgBox DecToHex(qwe), vbInformation, "Dec To Hex"
  End If
  Exit Sub
ErrHnd:
  Resume Debut
End Sub

Private Sub Command2_Click()
  Dim qaz As String
  Do
    qaz = InputBox("Enter a Hex : 80000000 to 7FFFFFFF", "Hex To Dec", 0)
  Loop Until Len(qaz) <= 8
  If qaz <> "" Then
    MsgBox HexToDec(qaz), vbInformation, "Hex To Dec"
  End If
End Sub

Private Sub Command3_Click()
  Dim qaz As String
  Do
    qaz = InputBox("Enter a Bin :" & vbCrLf & "10000000000000000000000000000000   to" & vbCrLf & "01111111111111111111111111111111", "Bin To Dec", 0)
  Loop Until Len(qaz) <= 32
  If qaz <> "" Then
    MsgBox BinToDec(qaz), vbInformation, "Bin To Dec"
  End If
End Sub

Private Sub Command4_Click()
  On Local Error GoTo ErrHnd
  Dim qwe As Long, qaz As String
Debut:
  qaz = InputBox("Enter a Long : -2147483648 to 2147483647", "Dec To Bin", 0)
  If qaz <> "" Then
    qwe = qaz
    MsgBox DecToBin(qwe), vbInformation, "Dec To Bin"
  End If
  Exit Sub
ErrHnd:
  Resume Debut
End Sub

Private Sub Command5_Click()
  Dim qaz As String
  Do
    qaz = InputBox("Enter a Hex : 80000000 to 7FFFFFFF", "Hex To Bin", 0)
  Loop Until Len(qaz) <= 8
  If qaz <> "" Then
    MsgBox HexToBin(qaz), vbInformation, "Hex To Bin"
  End If
End Sub

Private Sub Command6_Click()
  Dim qaz As String
  Do
    qaz = InputBox("Enter a Bin :" & vbCrLf & "10000000000000000000000000000000   to" & vbCrLf & "01111111111111111111111111111111", "Bin To Hex", 0)
  Loop Until Len(qaz) <= 32
  If qaz <> "" Then
    MsgBox BinToHex(qaz), vbInformation, "Bin To Hex"
  End If
End Sub

Private Sub Command7_Click()
  On Local Error GoTo ErrHnd
  Dim qwe As Integer, qaz As String
Debut:
  qaz = InputBox("Enter a Integer : -32768 to 32767", "High Byte & Low Byte", 0)
  If qaz <> "" Then
    qwe = qaz
    MsgBox "High Byte : " & HiByte(qwe) & vbCrLf & "Low Byte : " & LoByte(qwe), vbInformation, "High Byte & Low Byte"
  End If
  Exit Sub
ErrHnd:
  Resume Debut
End Sub

Private Sub Command8_Click()
  On Local Error GoTo ErrHnd
  Dim qwe As Long, qaz As String
Debut:
  qaz = InputBox("Enter a Long : -2147483648 to 2147483647", "High Word & Low Word", 0)
  If qaz <> "" Then
    qwe = qaz
    MsgBox "High Word : " & HiWord(qwe) & vbCrLf & vbCrLf _
           & "      High Byte : " & HiByteHiWord(qwe) & vbCrLf _
           & "      Low Byte : " & LoByteHiWord(qwe) & vbCrLf & vbCrLf & vbCrLf _
           & "Low Word : " & LoWord(qwe) & vbCrLf & vbCrLf _
           & "      High Byte : " & HiByteLoWord(qwe) & vbCrLf _
           & "      Low Byte : " & LoByteLoWord(qwe), _
           vbInformation, "High Word & Low Word"
  End If
  Exit Sub
ErrHnd:
  Resume Debut
End Sub

Private Sub Command9_Click()
  Swap var1, var2
  Text1.Text = var1
  Text2.Text = var2
End Sub

Private Sub Form_Load()
  var1 = 1.2
  var2 = 3.4
  Text1.Text = var1
  Text2.Text = var2
End Sub
