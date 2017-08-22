VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrCAM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'FOR WEBCAM DECLARATIONS
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Private mCapHwnd As Long

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054

Public Sub STOPCAM()
    DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
End Sub

Public Sub STARTCAM(HW As Long, H As Long, W As Long)
    mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, W, H, HW, 0)
    DoEvents
    SendMessage mCapHwnd, CONNECT, 0, 0
End Sub

Public Sub upFoto()
    'Get the picture from camera.. the main part
    SendMessage mCapHwnd, GET_FRAME, 0, 0
    SendMessage mCapHwnd, COPY, 0, 0
    Form1.Picture1.Picture = Clipboard.GetData
    Clipboard.Clear
End Sub
