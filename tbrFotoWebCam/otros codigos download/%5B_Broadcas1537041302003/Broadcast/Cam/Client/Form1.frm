VERSION 5.00
Object = "{95A385DC-B15E-4285-9F45-49F3B6DEABA6}#1.0#0"; "AVPhone3.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5928
   ClientLeft      =   4440
   ClientTop       =   2040
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5928
   ScaleWidth      =   5520
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Height          =   336
      Left            =   2052
      TabIndex        =   1
      Top             =   180
      Width           =   1596
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   336
      Left            =   288
      TabIndex        =   0
      Top             =   180
      Width           =   1596
   End
   Begin AVPhone3.UDPSocket UDPSocket1 
      Left            =   1152
      Top             =   2916
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0000
   End
   Begin AVPhone3.AudRnd AudRnd1 
      Left            =   2340
      Top             =   3060
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0024
   End
   Begin AVPhone3.VidRnd VidRnd1 
      Height          =   1848
      Left            =   324
      Top             =   684
      Width           =   1668
      _ExtentX        =   2942
      _ExtentY        =   3260
      Control         =   "Form1.frx":0048
   End
   Begin VB.Menu mnuShowCode 
      Caption         =   "&Show code!"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub Command1_Click()
    On Error GoTo ErrorHandle
    
    'set default dest address to user
    'entered
    UDPSocket1.SetSendAddress InputBox("Enter remote name or IP:", "Connect to server", UDPSocket1.GetIP(UDPSocket1.LocalAddress))
    
    'send out a connect message
    UDPSocket1.Frame 0, TM_CONNECT
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub Disconnect()
    'tell server we need disconnect
    UDPSocket1.Frame 0, TM_DISCONNECT
    
    StopRender
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrorHandle
    Disconnect
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub Form_Load()
    Caption = "Broadcast client"
    
    'bind to different port so that you
    'can try both server and client app
    'on the same machine.
    UDPSocket1.Bind 1721, 1720
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    Disconnect
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub StopRender()
    'set default dest address to 0
    UDPSocket1.SetSendAddress 0
    
    'stop video and audio
    AudRnd1.Format = vbNullString
    VidRnd1.Format = vbNullString
End Sub

Private Sub mnuShowCode_Click()
    On Error GoTo ErrorHandle
    ShowCode "..\..\", "form1.frm", "..\..\modmsgdef.bas", "..\WebClient\videowindow.ctl"
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub UDPSocket1_Frame(ByVal Address As Long, ByVal Handle As Long, ByVal lParam As Long, vData As Variant)
    On Error GoTo ErrorHandle
    Select Case Handle
    Case TM_DISCONNECT
    
        'set request a disconnecting
        StopRender
        
    Case TM_VIDEOFORMAT
        'video format
        VidRnd1.Format = vData
    Case TM_AUDIOFORMAT
        'audio format
        AudRnd1.Format = vData
    Case TM_VIDEORATE
        'video speed
        VidRnd1.Rate = vData
    
    Case TM_VIDEOFRAME
        'normal video frame
        VidRnd1.Frame vData, False
    Case TM_VIDEOFRAMEKEY
        'key video frame
        VidRnd1.Frame vData, True
        
    Case TM_AUDIOFRAME
        'audio frame
        AudRnd1.Frame vData
        
    End Select
    Exit Sub
    
ErrorHandle:
End Sub
