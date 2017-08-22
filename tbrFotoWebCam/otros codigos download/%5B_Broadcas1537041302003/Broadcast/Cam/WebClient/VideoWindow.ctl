VERSION 5.00
Object = "{95A385DC-B15E-4285-9F45-49F3B6DEABA6}#1.0#0"; "AVPhone3.ocx"
Begin VB.UserControl VideoWindow 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Begin AVPhone3.UDPSocket UDPSocket1 
      Left            =   2196
      Top             =   1872
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "VideoWindow.ctx":0000
   End
   Begin AVPhone3.AudRnd AudRnd1 
      Left            =   504
      Top             =   1944
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "VideoWindow.ctx":0024
   End
   Begin AVPhone3.VidRnd VidRnd1 
      Height          =   1380
      Left            =   468
      Top             =   324
      Width           =   1776
      _ExtentX        =   3112
      _ExtentY        =   2434
      Control         =   "VideoWindow.ctx":0048
   End
End
Attribute VB_Name = "VideoWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'==========================================================================
'  This is a part of Banasoft AVPhone controls
'  To get the last version of the control, please visit:
'
'  http://www.banasoft.net/AVPhone.htm
'
'  THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
'  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
'  PURPOSE.
'
'  Copyright (c) - 2002  Banasoft.  All Rights Reserved.
'
'==========================================================================


'public interface enable web script connect
Public Sub Connect(Host As String)

    'set default dest to host
    UDPSocket1.SetSendAddress Host
    
    'tell server we need connect
    UDPSocket1.Frame 0, TM_CONNECT
End Sub


Public Sub Disconnect()
    'tell server we need disconnect
    UDPSocket1.Frame 0, TM_DISCONNECT

    StopRender
End Sub


Private Sub ShowErr()
    MsgBox Err.Description, vbCritical
End Sub

Private Sub UserControl_Initialize()
    On Error GoTo ErrorHandle
    'move vidrnd to left-top
    VidRnd1.Move 0, 0
    
    'bind to port
    UDPSocket1.Bind 1721, 1720
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub UserControl_Resize()
    On Error GoTo ErrorHandle
    SizeMe
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo ErrorHandle
    Disconnect
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub StopRender()
    'set default dest to 0
    UDPSocket1.SetSendAddress 0
    
    'stop audio and video
    AudRnd1.Format = vbNullString
    VidRnd1.Format = vbNullString
End Sub

Private Sub UDPSocket1_Frame(ByVal Address As Long, ByVal Handle As Long, ByVal lParam As Long, vData As Variant)
    On Error GoTo ErrorHandle
    Select Case Handle
    Case TM_DISCONNECT
        'server request a disconnect
        StopRender
        
    Case TM_VIDEOFORMAT
        'video format
        VidRnd1.Format = vData
    Case TM_VIDEORATE
        'video speed
        VidRnd1.Rate = vData
    Case TM_AUDIOFORMAT
        'audio format
        AudRnd1.Format = vData
        
    Case TM_VIDEOFRAME
        'video frame normal
        VidRnd1.Frame vData, False
    Case TM_VIDEOFRAMEKEY
        'video frame key
        VidRnd1.Frame vData, True
        
    Case TM_AUDIOFRAME
        'audio frame
        AudRnd1.Frame vData
        
    End Select
    Exit Sub
    
ErrorHandle:
End Sub


Private Sub SizeMe()
    'size me to vidrnd control
    Size VidRnd1.Width, VidRnd1.Height
End Sub

Private Sub VidRnd1_Click()
    On Error GoTo ErrorHandle
    If VidRnd1.Zoom = -1 Then
        'restore window sizze
        VidRnd1.Zoom = 100
    Else
        'rendering in full screen
        VidRnd1.Zoom = -1
    End If
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub VidRnd1_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    'restore while user press "ESC"
    If KeyAscii = vbKeyEscape Then If VidRnd1.Zoom = -1 Then VidRnd1.Zoom = 100
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub VidRnd1_Resize()
    On Error GoTo ErrorHandle
    SizeMe
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub
