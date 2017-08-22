VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   1320
   ClientTop       =   1395
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9270
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   5295
      Left            =   3000
      ScaleHeight     =   5235
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   360
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   5520
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      Filter          =   "*.png|*.png"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

