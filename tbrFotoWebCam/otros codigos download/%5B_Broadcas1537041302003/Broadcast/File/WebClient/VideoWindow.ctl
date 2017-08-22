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
      Left            =   1692
      Top             =   1908
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "VideoWindow.ctx":0000
   End
   Begin AVPhone3.AudRnd AudRnd1 
      Left            =   540
      Top             =   1872
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "VideoWindow.ctx":0024
   End
   Begin AVPhone3.VidRnd VidRnd1 
      Height          =   1344
      Left            =   180
      Top             =   180
      Width           =   1596
      _ExtentX        =   2815
      _ExtentY        =   2371
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

'playing flag
Private blnPlaying As Boolean

Private Sub UserControl_Initialize()

    VidRnd1.Move 0, 0
    
    'bin to different port for enable local
    'loop testing
    UDPSocket1.Bind 1721, 1720
End Sub


Public Sub SetHost(Host As String)
    
    If blnPlaying Then StopFile
    
    'set default dest address to host
    Dim l As Long
    l = UDPSocket1.SetSendAddress(Host)
    
    'in this control we haven't any UI
    'we needn't tell server we need file list
    'If l Then UDPSocket1.Frame 0, TM_DIRECTORYINFO
    
End Sub

Public Sub PlayFile(Path As String)
    If blnPlaying Then StopFile
    
    'tell server we need play the file
    UDPSocket1.Frame 0, TM_CONNECT, , Path
End Sub

Public Sub StopFile()
    blnPlaying = False
    
    'tell server we need stop current file
    UDPSocket1.Frame 0, TM_DISCONNECT
    
    StopRender
End Sub

Private Sub StopRender()

    'stop video and audio
    VidRnd1.Format = vbNullString
    AudRnd1.Format = vbNullString
End Sub


Private Sub SizeMe()
    Size VidRnd1.Width, VidRnd1.Height
End Sub


Private Sub ShowErr()
    MsgBox Err.Description, vbCritical
End Sub

Private Sub UserControl_Resize()
    On Error GoTo ErrorHandle
    SizeMe
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub UserControl_Terminate()
    If blnPlaying Then StopFile
End Sub


Private Sub UDPSocket1_Frame(ByVal Address As Long, ByVal Handle As Long, ByVal Param As Long, Data As Variant)
    On Error GoTo ErrorHandle
    Select Case Handle
    Case TM_DIRECTORYINFO
        'we don't cognize this message
        'ListFiles Data
        
    Case TM_CONNECT
    
        'file opened
        blnPlaying = True
        
    Case Else
    
        If Not blnPlaying Then Exit Sub
        
        Select Case Handle
        Case TM_DISCONNECT
            'server stopped the file playing
            blnPlaying = False
            StopRender
        
        Case TM_AUDIOFORMAT
            'audio format
            AudRnd1.Format = Data
        Case TM_VIDEOFORMAT
            'video format
            VidRnd1.Format = Data
        Case TM_VIDEORATE
            'video rate
            VidRnd1.Rate = Data
            
        Case TM_AUDIOFRAME
            'audio frames
            AudRnd1.Frame Data
        Case TM_VIDEOFRAME
            'video normal frames
            VidRnd1.Frame Data, False
        Case TM_VIDEOFRAMEKEY
            'video key frames
            VidRnd1.Frame Data, True
                
        Case TM_MESSAGE
            
            'server error returned
            Select Case Param
            Case &H8004406D
            
                'at the end of the file
                StopFile
                
            Case Else
            
                'instead of showing the error
                'control should not have it 's own UI
                'let end user to determin
                'how to deal with the error
                'you may raise a event to notify
                'the end user
                'RaiseEvent Error(Param, Data)
                'here just for a test use msgbox
                '
                'StatuMsg "Error: " & Param & ", " & Data
                '
                'use a boolean flag avoid reenter
                Static b As Boolean
                If Not b Then
                    b = True
                    'MsgBox "Error: " & Param & ", " & Data, vbExclamation
                    b = False
                End If
                
            End Select
            
        End Select
    End Select
    Exit Sub
    
ErrorHandle:
    'ShowErr
End Sub


Private Sub VidRnd1_BufferEmpty()
    On Error GoTo ErrorHandle
    'request new video frame
    If blnPlaying Then UDPSocket1.Frame 0, TM_VIDEOFRAME
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub AudRnd1_BufferEmpty()
    On Error GoTo ErrorHandle
    'request new audio frame
    If blnPlaying Then UDPSocket1.Frame 0, TM_AUDIOFRAME
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub VidRnd1_Click()
    On Error GoTo ErrorHandle
    'switch full screen
    VidRnd1.Zoom = IIf(VidRnd1.Zoom = -1, 100, -1)
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub VidRnd1_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandle
    'restore while "ESC" pressed
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
