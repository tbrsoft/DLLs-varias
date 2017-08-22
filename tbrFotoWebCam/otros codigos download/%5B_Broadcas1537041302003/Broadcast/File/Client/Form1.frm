VERSION 5.00
Object = "{95A385DC-B15E-4285-9F45-49F3B6DEABA6}#1.0#0"; "AVPhone3.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5928
   ClientLeft      =   924
   ClientTop       =   2040
   ClientWidth     =   9912
   LinkTopic       =   "Form1"
   ScaleHeight     =   5928
   ScaleWidth      =   9912
   Begin VB.CommandButton Command1 
      Caption         =   "&List Server"
      Height          =   336
      Left            =   72
      TabIndex        =   0
      Top             =   180
      Width           =   1848
   End
   Begin VB.ListBox List1 
      Height          =   4188
      Left            =   72
      TabIndex        =   1
      Top             =   1008
      Width           =   1848
   End
   Begin AVPhone3.UDPSocket UDPSocket1 
      Left            =   2916
      Top             =   3528
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0000
   End
   Begin AVPhone3.AudRnd AudRnd1 
      Left            =   2196
      Top             =   3492
      _ExtentX        =   677
      _ExtentY        =   677
      Control         =   "Form1.frx":0024
   End
   Begin AVPhone3.VidRnd VidRnd1 
      Height          =   2964
      Left            =   2016
      Top             =   180
      Width           =   3324
      _ExtentX        =   5863
      _ExtentY        =   5228
      Control         =   "Form1.frx":0048
   End
   Begin VB.Label Label2 
      Caption         =   "Click for playing:"
      Height          =   444
      Left            =   72
      TabIndex        =   3
      Top             =   720
      Width           =   1812
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   300
      Left            =   108
      TabIndex        =   2
      Top             =   5544
      Width           =   5988
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

'play flag
Private blnPlaying As Boolean


Private Sub Form_Load()
    Caption = "Broadcast file client"
    
    'init to a statubar
    Label1 = vbNullString
    Label1.BorderStyle = 1
    
    'bin to different port for enable local
    'loop testing
    UDPSocket1.Bind 1721, 1720
    
    Show
    
    'click then "list" button
    SendKeys " "
End Sub


Private Sub StatuMsg(Msg As String)

    'beep
    Beep
    
    'show the msg
    Label1 = Msg
End Sub


Private Sub Command1_Click()

    With UDPSocket1
        Dim s As String
        s = InputBox("Enter server name or IP:", "Connect to server", .GetIP(.LocalAddress))
    End With
    If Len(s) <= 0 Then Exit Sub
    
    List1.Clear
    
    SetHost s
End Sub


Private Sub List1_Click()
    On Error GoTo ErrorHandle
    With List1
    
        'get the item
        Dim s As String
        s = .List(.ListIndex)
        
        'play it
        PlayFile .Tag & s
    End With
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub ListFiles(Path As Variant)
    List1.Clear
    
    'path returned from server include
    'file names splited by CRLF
    Dim l As Long
    Dim s As String
    Do
    
        l = InStr(Path, vbCrLf)
        If l <= 0 Then Exit Do
        
        s = Left$(Path, l - 1)
        If Len(s) > 0 Then List1.AddItem s
        Path = Mid$(Path, l + Len(vbCrLf))
    Loop
    
    s = Path
    If Len(s) > 0 Then List1.AddItem s
End Sub


Private Sub SetHost(Host As String)
    
    If blnPlaying Then StopFile
    
    'set default dest address to host
    Dim l As Long
    l = UDPSocket1.SetSendAddress(Host)
    
    'tell server we need file list
    If l Then UDPSocket1.Frame 0, TM_DIRECTORYINFO
    
End Sub


Private Sub PlayFile(Path As String)
    If blnPlaying Then StopFile
    
    'tell server we need play the file
    UDPSocket1.Frame 0, TM_CONNECT, , Path
End Sub

Private Sub StopFile()
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

Private Sub Form_Unload(Cancel As Integer)
    If blnPlaying Then StopFile
End Sub


Private Sub mnuShowCode_Click()
    On Error GoTo ErrorHandle
    ShowCode "..\..\", "form1.frm", "..\..\modmsgdef.bas", "..\WebClient\videowindow.ctl"
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub UDPSocket1_Frame(ByVal Address As Long, ByVal Handle As Long, ByVal Param As Long, Data As Variant)
    On Error GoTo ErrorHandle
    Select Case Handle
    Case TM_DIRECTORYINFO
        'file list returned
        ListFiles Data
        
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
            
                'show the error
                StatuMsg "Remote Error: " & Param & ", " & Data
            End Select
            
        End Select
    End Select
    Exit Sub
    
ErrorHandle:
    'show the error
    StatuMsg "Local Error: " & Err & ", " & Err.Description
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
