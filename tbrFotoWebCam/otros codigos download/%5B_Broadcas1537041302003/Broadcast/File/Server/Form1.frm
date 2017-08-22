VERSION 5.00
Object = "{95A385DC-B15E-4285-9F45-49F3B6DEABA6}#1.0#0"; "AVPhone3.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   6315
   ClientTop       =   2040
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   5520
   Begin VB.CommandButton Command2 
      Caption         =   "AVI &Directory..."
      Height          =   372
      Left            =   72
      TabIndex        =   2
      Top             =   2592
      Width           =   1740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Remove"
      Height          =   372
      Left            =   3744
      TabIndex        =   1
      Top             =   216
      Width           =   1560
   End
   Begin VB.ListBox List1 
      Height          =   4380
      Left            =   1980
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   792
      Width           =   3360
   End
   Begin AVPhone3.UDPSocket UDPSocket1 
      Left            =   468
      Top             =   936
      _ExtentX        =   847
      _ExtentY        =   847
      Control         =   "Form1.frx":0000
   End
   Begin AVPhone3.AVIFile AVIFile1 
      Index           =   0
      Left            =   468
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      Control         =   "Form1.frx":0024
   End
   Begin VB.Label Label2 
      Caption         =   "Active clients:"
      Height          =   300
      Left            =   1980
      TabIndex        =   4
      Top             =   288
      Width           =   2748
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   300
      Left            =   144
      TabIndex        =   3
      Top             =   5544
      Width           =   5232
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

'clients collection
Private clClients As Collection
'index collection for avifile controls
Private clAVIs As Collection

'server root path for all avi files.
Private sRootPath As String


Private Sub Form_Load()
    'A "poll" way broadcast demo
    Caption = "Broadcast File demo"
    
    'create the collection
    Set clClients = New Collection
    Set clAVIs = New Collection
    
    'init label to a statubar
    Label1.Caption = vbNullString
    Label1.BorderStyle = 1
    
    'bind to different ports to enable
    'local machine loop back testing.
    With UDPSocket1
        .Bind 1720, 1721
        
        'show my IP in title bar
        Caption = Caption & " " & .GetIP(.LocalAddress)
    End With
    
    'set default path
    Dim s As String
    s = App.Path
    ChangeDirectory s

End Sub


Private Sub Command2_Click()
    On Error GoTo ErrorHandle
    
    'change default path
    Dim s As String
    s = OpenFileDlg(hwnd, 0)
    If Len(s) <= 0 Then Exit Sub
    
    'get the file path
    Dim lo As Long
    lo = InStr(s, ":")
    Do
        Dim l As Long
        l = InStr(lo + 1, s, "\")
        If l <= 0 Then Exit Do
        
        lo = l
    Loop
    s = Left$(s, lo)
    
    'changing it
    ChangeDirectory s
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub ChangeDirectory(Path As String)

    'fix path string
    If Right$(Path, 1) <> "\" Then Path = Path & "\"

    'store it to local variable
    sRootPath = Path
    
    'tell all clients path changed
    SendToClients TM_DIRECTORYINFO, , GetFileList(sRootPath)
End Sub


Private Sub Command1_Click()
    On Error GoTo ErrorHandle
    
    'get the list item
    Dim s As String
    With List1
        s = .List(.ListIndex)
    End With
    
    Dim l As Long
    With UDPSocket1
    
        'got address from ip
        l = .SetSendAddress(s)
        
        'remove the avi resource for the address
        CloseAVI l
        
        'tell the client we dropped it
        .Frame 0, TM_DISCONNECT
    End With
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub


Private Sub StatuError()
    'beep alarm
    Beep
    
    'show the error info
    Label1 = "Error: " & Err & ", " & Err.Description
End Sub


Private Function GetFileList(Path As String)

    'list all *.avi file in this directory
    Dim s As String
    s = Dir$(Path & "*.avi", vbDirectory)
    
    'combine to string by a CRLF split
    Do Until Len(s) <= 0
        GetFileList = GetFileList & s & vbCrLf
        
        'next file
        s = Dir$()
    Loop

    'list all *.wav file
    s = Dir$(Path & "*.wav", vbDirectory)
    
    Do Until Len(s) <= 0
        GetFileList = GetFileList & s & vbCrLf
        
        'next file
        s = Dir$()
    Loop
End Function


'send message to all of the client current registered
Private Sub SendToClients(ByVal Msg As Long, Optional ByVal lParam As Long, Optional Data As Variant)
    With UDPSocket1
        Dim v As Variant
        For Each v In clClients
            .Frame v, Msg, lParam, Data
        Next
    End With
End Sub


Private Sub ReturnError(ByVal Address As Long)
    'tell the user what error was happened
    UDPSocket1.Frame Address, TM_MESSAGE, Err, Err.Description
End Sub


Private Sub ListFile(ByVal Address As Long, Path As Variant)
    On Error GoTo ErrorHandle
    
    'remote user can't request to list files doesn't local on the root directory.
    If InStr(Path, "..\") Then Err.Raise 5, , "Invalid path"
    
    'tell the user all files
    UDPSocket1.Frame Address, TM_DIRECTORYINFO, , GetFileList(sRootPath & Path)
    Exit Sub
    
ErrorHandle:
    'any message caused error should return to the client
    ReturnError Address
End Sub


Private Sub OpenFile(ByVal Address As Long, Path As Variant)
    On Error GoTo ErrorHandle

    Dim s As String
    s = CStr(Address)
    
    'add to client collection
    clClients.Add Address, s
    
    'find a unused control index
    Dim l As Long
    Do
        l = l + 1
        
        Dim a As AVIFile
        For Each a In AVIFile1
            If a.Index = l Then Exit For
        Next
    Loop Until a Is Nothing
    
    'load a new avifile control for the address
    Load AVIFile1(l)
    
    'add to collection
    clAVIs.Add l, s
    
    'tell client we connected
    With UDPSocket1
        .Frame Address, TM_CONNECT
        
        'get client's IP
        s = .GetIP(Address)
    End With
    
    'add to list
    With List1
        .AddItem s
        .ItemData(.ListCount - 1) = Address
    End With
    
    With AVIFile1(l)
    
        'open the file on root path
        .OpenFile sRootPath & Path
        
        'get first audio track format
        Dim v1 As Variant
        v1 = .StreamFormat(-1)
        
        'get first video track format
        Dim v2 As Variant
        v2 = .StreamFormat(-2)
        
        'get first video track speed
        Dim v3 As Variant
        v3 = .StreamRate(-2)
    End With
    
    With UDPSocket1
        'if there is a valid audio track tell the format the client
        If Not IsNull(v1) Then .Frame Address, TM_AUDIOFORMAT, , v1
        
        'if there is a valid video track tell the format and rate to the client
        If Not IsNull(v2) Then
            .Frame Address, TM_VIDEORATE, , v3
            .Frame Address, TM_VIDEOFORMAT, , v2
        End If
    End With
    Exit Sub
    
ErrorHandle:
    ReturnError Address
End Sub


Private Sub CloseAVI(ByVal Address As Long)
    
    Dim s As String
    s = CStr(Address)
    
    'get related avifile index
    Dim l As Long
    l = clAVIs.Item(s)
    
    'unload the control
    Unload AVIFile1(l)
    
    'remove the client from our collections
    clClients.Remove s
    clAVIs.Remove s
    
    'remove the client IP from list
    With List1
        For l = 0 To .ListCount - 1
            If .ItemData(l) = Address Then Exit For
        Next
        .RemoveItem l
    End With
    
End Sub


Private Sub CloseFile(ByVal Address As Long)
    On Error GoTo ErrorHandle
    'close related controls
    CloseAVI Address
    Exit Sub
    
ErrorHandle:
    ReturnError Address
End Sub


Private Sub GetFrame(ByVal Address As Long, ByVal Track As Long)

    'get related avifile control's index
    Dim l As Long
    l = clAVIs.Item(CStr(Address))
    
    On Error GoTo ErrorHandle
    
    'read data
    Dim bt() As Byte
    Dim b As Boolean
    AVIFile1(l).StreamRead Track, bt, b
    
    If Track = -2 Then
        'video track, check if it is a key frame
        l = IIf(b, TM_VIDEOFRAMEKEY, TM_VIDEOFRAME)
    Else
        'audio track
        l = TM_AUDIOFRAME
    End If
    
    'send to the client
    UDPSocket1.Frame Address, l, , bt
    Exit Sub
    
ErrorHandle:
    ReturnError Address
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    'tell all clients we are exiting
    SendToClients TM_DISCONNECT
    Exit Sub
    
ErrorHandle:
    StatuError
    Resume Next
End Sub

Private Sub mnuShowCode_Click()
    On Error GoTo ErrorHandle
    ShowCode "..\..\", "form1.frm", "..\..\modmsgdef.bas", "..\..\..\..\module1.bas"
    Exit Sub
    
ErrorHandle:
    ShowErr
End Sub

Private Sub UDPSocket1_Frame(ByVal Address As Long, ByVal Handle As Long, ByVal Param As Long, Data As Variant)
    On Error GoTo ErrorHandle
    Select Case Handle
    Case TM_DIRECTORYINFO
        'need list files
        ListFile Address, Data
        
    Case TM_CONNECT
        'need play a file
        OpenFile Address, Data
    
    Case TM_DISCONNECT
        'need stop the playing
        CloseFile Address
        
    Case TM_VIDEOFRAME
        'need new video frame
        GetFrame Address, -2
    
    Case TM_AUDIOFRAME
        'need new audio frame
        GetFrame Address, -1
        
    End Select
    Exit Sub
    
ErrorHandle:
    StatuError
End Sub
