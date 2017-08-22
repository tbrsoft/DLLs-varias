VERSION 5.00
Object = "{95A385DC-B15E-4285-9F45-49F3B6DEABA6}#1.0#0"; "AVPhone3.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5916
   ClientLeft      =   888
   ClientTop       =   2100
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5916
   ScaleWidth      =   5520
   Begin VB.CommandButton Command6 
      Caption         =   "&Remove"
      Height          =   336
      Left            =   576
      TabIndex        =   8
      Top             =   5040
      Width           =   1704
   End
   Begin VB.ListBox List1 
      Height          =   4188
      Left            =   2808
      TabIndex        =   7
      Top             =   828
      Width           =   2496
   End
   Begin VB.Frame Frame2 
      Caption         =   "Audio"
      Height          =   1452
      Left            =   468
      TabIndex        =   4
      Top             =   3060
      Width           =   1884
      Begin VB.CommandButton Command5 
         Caption         =   "For&mat"
         Height          =   300
         Left            =   144
         TabIndex        =   6
         Top             =   324
         Width           =   1524
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Co&dec"
         Height          =   336
         Left            =   144
         TabIndex        =   5
         Top             =   792
         Width           =   1524
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Video"
      Height          =   1992
      Left            =   432
      TabIndex        =   0
      Top             =   324
      Width           =   1956
      Begin VB.CommandButton Command3 
         Caption         =   "&Format"
         Height          =   300
         Left            =   180
         TabIndex        =   3
         Top             =   972
         Width           =   1524
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Codec"
         Height          =   336
         Left            =   216
         TabIndex        =   2
         Top             =   1440
         Width           =   1488
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Source"
         Height          =   336
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   1524
      End
   End
   Begin AVPhone3.UDPSocket UDPSocket1 
      Left            =   1692
      Top             =   2448
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0000
   End
   Begin AVPhone3.AudCodec AudCodec1 
      Left            =   1620
      Top             =   2772
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0024
   End
   Begin AVPhone3.VidCodec VidCodec1 
      Left            =   2088
      Top             =   2700
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0048
   End
   Begin AVPhone3.AudCap AudCap1 
      Left            =   504
      Top             =   2664
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":006C
   End
   Begin AVPhone3.VidCap VidCap1 
      Left            =   1152
      Top             =   2664
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0090
   End
   Begin VB.Label Label2 
      Caption         =   "&Active Clients:"
      Height          =   300
      Left            =   2808
      TabIndex        =   10
      Top             =   432
      Width           =   2496
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   336
      Left            =   108
      TabIndex        =   9
      Top             =   5508
      Width           =   5268
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

'client address collection
Private clClients As Collection


Private Sub Form_Load()
    'A "push" way broadcast demo"
    Caption = "Broadcast server"
    
    'init label as a statubar
    Label1.Caption = vbNullString
    Label1.BorderStyle = 1
    
    'create the colloection, use this both this
    'collection and listbox control to store clients
    'because memory collection is much fast than
    'UI controls
    Set clClients = New Collection
    
    'bind to different port so that you
    'can try both server and client app
    'on the same machine.
    UDPSocket1.Bind 1720, 1721
    
    'show local IP on caption
    Caption = Caption & " " & UDPSocket1.GetIP(UDPSocket1.LocalAddress)
    
    'show me
    Show
    
    'connect to video device if it is available
    ConnectVideo
    
    'connect to audio device if it is available
    ConnectAudio
End Sub


'show error message on statubar instead of msgbox
'because normal no UI on server
Private Sub StatusErrMsg()
    'beep
    Beep
    
    'show error
    Label1 = "Error: " & Err & ", " & Err.Description
End Sub


'send message to all of the client current registered
Private Sub SendToClients(ByVal Msg As Long, Optional ByVal lParam As Long, Optional Data As Variant)
    With UDPSocket1
        Dim v As Variant
        For Each v In clClients
            .Frame v, Msg, lParam, Data
        Next
    End With
End Sub


'new audio device connected or format changed
Private Sub NewAudioFormat()

    'stop codec
    AudCodec1.InFormat = vbNullString
    
    'init to the codec you wanted.
    AudCodec1.OutFormat = "112,1,8000,16,600,12"
    
    'start compression
    AudCodec1.InFormat = AudCap1.Format
    
    'new format
    NewAudioCompFormat
End Sub

'tell the clients our audio format
Private Sub NewAudioCompFormat()
    SendToClients TM_AUDIOFORMAT, , AudCodec1.OutFormat
End Sub

'new video device connected or format changed
Private Sub NewVideoFormat()

    'stop codec
    VidCodec1.InFormat = vbNullString
    
    'set to codec and quality to you wanted
    VidCodec1.OutFormat = "iv50"
    VidCodec1.Quality = 40
    
    'start compression
    VidCodec1.InFormat = VidCap1.Format
    
    'new format
    NewVideoCompFormat
End Sub


'tell clients our video format
Private Sub NewVideoCompFormat()
    SendToClients TM_VIDEOFORMAT, , VidCodec1.OutFormat
End Sub

'video source dialog
Private Sub Command1_Click()
    On Error GoTo ErrorHandle
    VidCap1.SourceDlg
    
    'tell clients video format changed
    NewVideoFormat
    Exit Sub
    
ErrorHandle:
    If Err <> 32755 Then ShowErr
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrorHandle
    VidCodec1.CompressorDlg
    
    'tell clients video format changed
    NewVideoCompFormat
    Exit Sub
    
ErrorHandle:
    If Err <> 32755 Then ShowErr
End Sub

Private Sub Command3_Click()
    On Error GoTo ErrorHandle
    VidCap1.FormatDlg
    
    'tell clients video format changed
    NewVideoFormat
    Exit Sub
    
ErrorHandle:
    If Err <> 32755 Then ShowErr

End Sub

Private Sub Command4_Click()
    On Error GoTo ErrorHandle
    AudCodec1.CompressorDlg
    
    'tell clients audio format changed
    NewAudioCompFormat
    Exit Sub
    
ErrorHandle:
    If Err <> 32755 Then ShowErr
End Sub

Private Sub Command5_Click()
    On Error GoTo ErrorHandle
    AudCap1.FormatDlg
    
    'tell clients audio format changed
    NewAudioFormat
    Exit Sub
    
ErrorHandle:
    If Err <> 32755 Then ShowErr
End Sub


Private Sub RemoveClient(ByVal Address As Long)

    'remove it from collection
    clClients.Remove CStr(Address)
    
    'remove it from list
    With List1
        Dim l As Long
        For l = 0 To .ListCount - 1
            If .ItemData(l) = Address Then
                .RemoveItem l
                Exit For
            End If
        Next
    End With
End Sub


Private Sub Command6_Click()
    On Error GoTo ErrorHandle
    Dim l As Long
    With List1
        l = .ItemData(.ListIndex)
    End With
    
    'remove it
    RemoveClient l
    
    'tell the client we dropped it
    UDPSocket1.Frame l, TM_DISCONNECT
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub ConnectVideo()
    On Error GoTo ErrorHandle
    
    'connect to first available video device
    VidCap1.Device = -1
    
    'new video format connected
    NewVideoFormat
    Exit Sub
    
ErrorHandle:
    StatusErrMsg
End Sub


Private Sub ConnectAudio()
    On Error GoTo ErrorHandle
    'connect to default wavein device
    AudCap1.Device = -1
    
    'new audio format connected
    NewAudioFormat
    Exit Sub
    
ErrorHandle:
    StatusErrMsg
End Sub


Private Sub AudCap1_Frame(Data As Variant)
    On Error GoTo ErrorHandle
    
    'write to compressor
    AudCodec1.Frame Data
    Exit Sub
ErrorHandle:
    StatusErrMsg
End Sub

Private Sub AudCodec1_Frame(Data As Variant)
    On Error GoTo ErrorHandle
    
    'audio data compressed, write to all clients
    SendToClients TM_AUDIOFRAME, , Data
    Exit Sub
ErrorHandle:
    StatusErrMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    
    'tell clients we exit
    SendToClients TM_DISCONNECT
    Exit Sub
ErrorHandle:
    StatusErrMsg
End Sub

Private Sub mnuShowCode_Click()
    On Error GoTo ErrorHandle
    ShowCode "..\..\", "form1.frm", "..\..\modmsgdef.bas"
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub VidCap1_Frame(Data As Variant)
    On Error GoTo ErrorHandle
    
    'write to compression
    VidCodec1.Frame Data
    Exit Sub
ErrorHandle:
    StatusErrMsg
End Sub

Private Sub VidCodec1_Frame(Data As Variant, ByVal IsKeyFrame As Boolean)
    On Error GoTo ErrorHandle
    
    'video data compressed, write to all clients
    'with key frame indication
    SendToClients IIf(IsKeyFrame, TM_VIDEOFRAMEKEY, TM_VIDEOFRAME), , Data
    Exit Sub
ErrorHandle:
    StatusErrMsg
End Sub


'new client request to connect
Private Sub AddClient(ByVal Address As Long)

    'all to collection
    clClients.Add Address, CStr(Address)
    
    'tell the client our video and audio format
    With UDPSocket1
        .Frame Address, TM_VIDEOFORMAT, , VidCodec1.OutFormat
        .Frame Address, TM_AUDIOFORMAT, , AudCodec1.OutFormat
        
        'get it's IP
        Dim s As String
        s = .GetIP(Address)
    End With

    'add to list for showing
    With List1
        .AddItem s
        .ItemData(.ListCount - 1) = Address
    End With
End Sub


Private Sub UDPSocket1_Frame(ByVal Address As Long, ByVal Handle As Long, ByVal lParam As Long, vData As Variant)
    On Error GoTo ErrorHandle
    Select Case Handle
    Case TM_CONNECT
    
        'a client want to connect us
        AddClient Address
        'tell client our speed
        UDPSocket1.Frame Address, TM_VIDEORATE, , VidCap1.Rate
        
    Case TM_DISCONNECT
        'a client disconnected
        RemoveClient Address
    End Select
    Exit Sub
ErrorHandle:
    StatusErrMsg
End Sub


