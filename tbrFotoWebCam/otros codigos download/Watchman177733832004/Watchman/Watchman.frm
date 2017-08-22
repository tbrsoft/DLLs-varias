VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Watchman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Watchman"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   2370
   Icon            =   "Watchman.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   158
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider2 
      Height          =   180
      Left            =   -15
      TabIndex        =   3
      Top             =   4170
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   318
      _Version        =   393216
      Min             =   1
      Max             =   20
      SelStart        =   2
      Value           =   2
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   195
      Left            =   -30
      TabIndex        =   2
      Top             =   3495
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   344
      _Version        =   393216
      Min             =   1
      Max             =   20
      SelStart        =   5
      Value           =   5
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   -15
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2115
      Width           =   2400
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   -15
      TabIndex        =   0
      Top             =   2910
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Motion Counter"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   7
      Top             =   1830
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Motion Grade"
      Height          =   255
      Index           =   2
      Left            =   675
      TabIndex        =   6
      Top             =   2610
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Lower Sensetivity Threshold"
      Height          =   255
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   3870
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "Upper Sensetivity Threshold"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3225
      Width           =   2100
   End
End
Attribute VB_Name = "Watchman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lwndC As Long     ' Handle to the Capture Windows
Dim lNFrames As Long  ' Number of frames captured
Sub ResizeCaptureWindow(ByVal lwnd As Long)
    Dim CAPSTATUS As CAPSTATUS
    '// Get the capture window attributes .. width and height
    capGetStatus lwnd, VarPtr(CAPSTATUS), Len(CAPSTATUS)
    '// Resize the capture window to the capture sizes
    SetWindowPos lwnd, HWND_BOTTOM, 0, 0, CAPSTATUS.uiImageWidth, CAPSTATUS.uiImageHeight, SWP_NOMOVE Or SWP_NOZORDER
End Sub
Private Sub Form_Load()
    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim Caps As CAPDRIVERCAPS
    '//Create Capture Window
    capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
    lwndC = capCreateCaptureWindowA(lpszName, WS_CHILD Or WS_VISIBLE, 0, 0, 160, 120, Me.hwnd, 0)
    '// Connect the capture window to the driver
    capDriverConnect lwndC, 0
    '// Get the capabilities of the capture driver
    capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
    '// Set the video stream callback function
    capSetCallbackOnVideoStream lwndC, AddressOf MyVideoStreamCallback
    capSetCallbackOnFrame lwndC, AddressOf MyFrameCallback
    '// Show Video Format Dialog
    capDlgVideoFormat lwndC
    '// Set the preview rate in milliseconds
    capPreviewRate lwndC, 25
    '// Start previewing the image from the camera
    capPreview lwndC, True
    '// Resize the capture window to show the whole image
    ResizeCaptureWindow lwndC
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '// Disable all callbacks
    capSetCallbackOnError lwndC, vbNull
    capSetCallbackOnStatus lwndC, vbNull
    capSetCallbackOnYield lwndC, vbNull
    capSetCallbackOnFrame lwndC, vbNull
    capSetCallbackOnVideoStream lwndC, vbNull
    capSetCallbackOnWaveStream lwndC, vbNull
    capSetCallbackOnCapControl lwndC, vbNull
End Sub
Private Sub Slider1_Change()
  If Slider1.Value < Slider2.Value Then Slider2.Value = Slider1.Value
End Sub
Private Sub Slider2_Change()
  If Slider1.Value < Slider2.Value Then Slider1.Value = Slider2.Value
End Sub

