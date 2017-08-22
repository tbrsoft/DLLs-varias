VERSION 5.00
Object = "{02C6579B-57DA-4DA0-AB1F-6191A8F3098D}#1.0#0"; "GpCapture.ocx"
Begin VB.Form frmTest 
   Caption         =   "Capture DV"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   StartUpPosition =   2  'CenterScreen
   Begin GPCAPTURELib.GpCapture GpCapture1 
      Height          =   5415
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   9551
      _StockProps     =   0
   End
   Begin VB.CommandButton cmdCap 
      Caption         =   "Capture2"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "Capture"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "connect"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   1095
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCap_Click()
    GpCapture1.CaptureFile = "d:\qqq2.avi"
    GpCapture1.StartCaptureType2
End Sub

Private Sub cmdCapture_Click()
    GpCapture1.CaptureFile = "d:\qqq1.avi"
    GpCapture1.StartCaptureType1
End Sub

Private Sub cmdConnect_Click()
    Debug.Print GpCapture1.Connect
End Sub

Private Sub cmdPause_Click()
    GpCapture1.Pause
End Sub

Private Sub cmdPreview_Click()
    GpCapture1.Preview
End Sub

Private Sub cmdStop_Click()
    GpCapture1.Stop
End Sub
