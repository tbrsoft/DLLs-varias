VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Native VB TAPI Tester"
   ClientHeight    =   5205
   ClientLeft      =   630
   ClientTop       =   1710
   ClientWidth     =   6705
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   Begin VB.CommandButton cmdHangUp 
      Caption         =   "Hang Up"
      Height          =   555
      Left            =   4230
      TabIndex        =   13
      Top             =   3465
      Width           =   2355
   End
   Begin VB.ListBox lstCallProgress 
      Height          =   645
      Left            =   4230
      TabIndex        =   11
      Top             =   2565
      Width           =   2370
   End
   Begin VB.PictureBox picIcon 
      Height          =   660
      Left            =   5940
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   9
      Top             =   135
      Width           =   660
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   555
      Left            =   4230
      TabIndex        =   8
      Top             =   4545
      Width           =   2355
   End
   Begin VB.CommandButton cmdDialProps 
      Caption         =   "Dialing Properties..."
      Height          =   555
      Left            =   2160
      TabIndex        =   6
      Top             =   4545
      Width           =   1860
   End
   Begin VB.TextBox txtPhoneNumber 
      Height          =   330
      Left            =   4230
      TabIndex        =   5
      Text            =   "123-4567"
      Top             =   1215
      Width           =   2355
   End
   Begin VB.CommandButton cmdDial 
      Caption         =   "Dial..."
      Height          =   555
      Left            =   4230
      TabIndex        =   4
      Top             =   1620
      Width           =   2355
   End
   Begin VB.CommandButton cmdConfigDlg 
      Caption         =   "Line Config Dialog..."
      Height          =   555
      Left            =   135
      TabIndex        =   3
      Top             =   4545
      Width           =   1860
   End
   Begin VB.ListBox lstLineInfo 
      Height          =   3765
      Left            =   135
      TabIndex        =   2
      Top             =   525
      Width           =   3885
   End
   Begin VB.ComboBox cboLineSel 
      Height          =   300
      Left            =   2610
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   1425
   End
   Begin VB.Label lblProgress 
      Caption         =   "Call Progress:"
      Height          =   240
      Left            =   4275
      TabIndex        =   12
      Top             =   2250
      Width           =   2370
   End
   Begin VB.Label lblIcon 
      Caption         =   "Icon from TSP:"
      Height          =   195
      Left            =   4230
      TabIndex        =   10
      Top             =   225
      Width           =   1200
   End
   Begin VB.Label lblPhoneNumber 
      Caption         =   "Enter Phone Number to dial:"
      Height          =   240
      Left            =   4230
      TabIndex        =   7
      Top             =   855
      Width           =   2370
   End
   Begin VB.Label lblLineSel 
      Caption         =   "Select TAPI Line number:"
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   210
      Width           =   2370
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'****************************************************************
'*  VB file:   fMain.frm...
'*             VB TAPI Devcaps/Dialer app
'*
'*  created:        1999 by Ray Mercer
'*
'*  last modified:  8/25/99 by Ray Mercer (added comments)
'*
'*  Copyright (c) 1998 Ray Mercer.  All rights reserved.
'*  Latest version at http://i.am/shrinkwrapvb
'****************************************************************
Private Declare Function timeGetTime Lib "winmm.dll" () As Long 'for slowing down GUI only

'dimension a TAPI variable with events
Public WithEvents tapiline As CvbTAPILine
Attribute tapiline.VB_VarHelpID = -1

Private Sub Form_Load()
Dim success As Boolean
Dim line As Long

cmdHangUp.Enabled = False
Me.Show
Me.MousePointer = vbHourglass
Me.Refresh

'Initialize the TAPI class
Set tapiline = New CvbTAPILine
'Set the negotiate API version between 1.3 and 3.0
'This step is not necessary if default versions are OK (1.3 - 3.0)
'It's easy to visualize the high and low words in Hex
tapiline.LowAPI = &H10003 ' 1.3 = &H00010003
tapiline.HiAPI = &H30000 '  3.0 = &H00030000

'give user feedback and pause to let user read it
lstLineInfo.AddItem "Negotiating TAPI version..."
lstLineInfo.Refresh
Call Pause(1000, False)

'initialize and negotiate API versions for all lines
success = tapiline.Create

'List all available lines with these API caps in the listbox
If success Then
    For line = 0 To tapiline.numLines - 1
        tapiline.CurrentLineID = line
        If tapiline.NegotiatedAPIVersion Then
            cboLineSel.AddItem (line)
        End If
    Next
    'set the currently selected line (and trigger the click event)
    cboLineSel.ListIndex = 0
Else
    'give user feedback so they know their TAPI device doesn't support 2.1
    lstLineInfo.AddItem tapiline.ErrorString(tapiline.LastError)
    lstLineInfo.Refresh
    Call Pause(500, False)
    lstLineInfo.AddItem "Trying TAPI 1.3 - 1.4..."
    lstLineInfo.Refresh
    Call Pause(1000, False)
    
    'now re-negotiate lower version
    tapiline.LowAPI = &H10003 ' 1.3 = &H00010003
    tapiline.HiAPI = &H10004 '  1.4 = &H00010004
    success = tapiline.Create
    If success Then
        For line = 0 To tapiline.numLines - 1
            tapiline.CurrentLineID = line
            If tapiline.NegotiatedAPIVersion Then
                cboLineSel.AddItem (line)
            End If
        Next
        'set the currently selected line (and trigger the click event)
        cboLineSel.ListIndex = 0
    Else
        lstLineInfo.AddItem "Failed to negotiate TAPI version! <Critical Error!>"
    End If
End If



Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'class automatically deallocates itself when it is destroyed
End Sub

Private Sub cboLineSel_Click()

lstLineInfo.Clear
If cboLineSel.List(cboLineSel.ListIndex) <> "" Then
    'this sets the current line in the TAPIline class
    tapiline.CurrentLineID = cboLineSel.List(cboLineSel.ListIndex)
    
    'this section just prints out a lot of info about the selected Line
    lstLineInfo.AddItem "TAPI LINE: #" & tapiline.CurrentLineID
    lstLineInfo.AddItem "TAPI LINE NAME: " & tapiline.LineName
    lstLineInfo.AddItem "TAPI PROVIDER INFO: " & tapiline.ProviderInfo
    lstLineInfo.AddItem "TAPI SWITCH INFO: " & tapiline.SwitchInfo
    lstLineInfo.AddItem "Permanent Line ID: " & tapiline.PermanentLineID
    Select Case tapiline.StringFormat
        Case STRINGFORMAT_ASCII
            lstLineInfo.AddItem "String Format: STRINGFORMAT_ASCII"
        Case STRINGFORMAT_DBCS
            lstLineInfo.AddItem "String Format: STRINGFORMAT_DBCS"
        Case STRINGFORMAT_UNICODE
            lstLineInfo.AddItem "String Format: STRINGFORMAT_UNICODE"
        Case STRINGFORMAT_BINARY
            lstLineInfo.AddItem "String Format: STRINGFORMAT_BINARY"
        Case Else
    End Select
    lstLineInfo.AddItem "Number of addresses associated with this line: " & tapiline.numAddresses
    lstLineInfo.AddItem "Max data rate: " & tapiline.maxDataRate
    lstLineInfo.AddItem "Bearer Modes supported:"
    If LINEBEARERMODE_VOICE And tapiline.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_VOICE"
    If LINEBEARERMODE_SPEECH And tapiline.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_SPEECH"
    If LINEBEARERMODE_DATA And tapiline.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_DATA"
    If LINEBEARERMODE_ALTSPEECHDATA And tapiline.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_ALTSPEECHDATA"
    If LINEBEARERMODE_MULTIUSE And tapiline.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_MULTIUSE"
    If LINEBEARERMODE_NONCALLSIGNALING And tapiline.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_NONCALLSIGNALING"
    lstLineInfo.AddItem "Address Modes supported:"
    If tapiline.AddressModes And LINEADDRESSMODE_ADDRESSID Then lstLineInfo.AddItem vbTab & "LINEADDRESSMODE_ADDRESSID"
    If tapiline.AddressModes And LINEADDRESSMODE_DIALABLEADDR Then lstLineInfo.AddItem vbTab & "LINEADDRESSMODE_DIALABLEADDR"
    lstLineInfo.AddItem "Media Modes supported:"
    If LINEMEDIAMODE_ADSI And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_ADSI"
    If LINEMEDIAMODE_AUTOMATEDVOICE And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_AUTOMATEDVOICE"
    If LINEMEDIAMODE_DATAMODEM And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_DATAMODEM"
    If LINEMEDIAMODE_DIGITALDATA And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_DIGITALDATA"
    If LINEMEDIAMODE_G3FAX And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_G3FAX"
    If LINEMEDIAMODE_G4FAX And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_G4FAX"
    If LINEMEDIAMODE_INTERACTIVEVOICE And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_INTERACTIVEVOICE"
    If LINEMEDIAMODE_MIXED And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_MIXED"
    If LINEMEDIAMODE_TDD And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_TDD"
    If LINEMEDIAMODE_TELETEX And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_TELETEX"
    If LINEMEDIAMODE_TELEX And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_TELEX"
    If LINEMEDIAMODE_UNKNOWN And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_UNKNOWN"
    If LINEMEDIAMODE_VIDEOTEX And tapiline.mediamodes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_VIDEOTEX"
    lstLineInfo.AddItem "Line Tone Generation supported: " & CBool(tapiline.GenerateToneMaxNumFreq)
    If CBool(tapiline.GenerateToneMaxNumFreq) Then 'show if tone generation is supported
        If LINETONEMODE_BEEP And tapiline.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_BEEP"
        If LINETONEMODE_BILLING And tapiline.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_BILLING"
        If LINETONEMODE_BUSY And tapiline.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_BUSY"
        If LINETONEMODE_CUSTOM And tapiline.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_CUSTOM"
        If LINETONEMODE_RINGBACK And tapiline.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_RINGBACK"
    End If
    lstLineInfo.AddItem "Number of terminals for this line: " & tapiline.numTerminals
Else
    lstLineInfo.AddItem "<No Valid TAPI Line Selected>"
End If
'now we check to see if this line supports making voice calls
If tapiline.LineSupportsVoiceCalls Then
    'and enable/disable the call btn as appropriate
    cmdDial.Enabled = True
Else
    cmdDial.Enabled = False
End If

'paint the icon for this line
picIcon.AutoRedraw = True
If tapiline.PaintDevIcon(picIcon.hdc, 4, 4) Then
    'make icon persistent
    picIcon.Picture = picIcon.Image
Else
    'TSP contains no icon - just erase the old one
    picIcon.Picture = LoadPicture()
End If


End Sub

Private Sub cmdConfigDlg_Click()
    Dim success As Boolean
    success = tapiline.ConfigDialog(Me.hWnd)
    If success <> True Then
        MsgBox tapiline.ErrorString(tapiline.LastError)
    End If
End Sub

Private Sub cmdDial_Click()
Dim success As Boolean

cmdDial.Enabled = False
'give the user some progress indication
lstCallProgress.AddItem "Opening line: #" & tapiline.CurrentLineID
success = tapiline.OpenLine
If success <> True Then
    MsgBox "TAPI ERROR " & vbCrLf & tapiline.ErrorString(tapiline.LastError), vbCritical, App.Title
    cmdDial.Enabled = True
    Exit Sub
End If
lstCallProgress.AddItem "Preparing to dial: " & txtPhoneNumber.Text
success = tapiline.MakeCallAsynch(txtPhoneNumber.Text)
If success <> True Then
    Call tapiline.CloseLine
    MsgBox "Error #" & tapiline.LastError & vbCrLf & _
            tapiline.ErrorString(tapiline.LastError), vbInformation, App.Title
    cmdDial.Enabled = True
    Exit Sub
End If


End Sub

Private Sub cmdDialProps_Click()
    Dim success As Boolean
    success = tapiline.DialingPropertiesDialog(Me.hWnd, txtPhoneNumber)
    If success <> True Then
        MsgBox tapiline.ErrorString(tapiline.LastError)
    End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHangUp_Click()
Dim success As Boolean

lstCallProgress.AddItem "Requesting drop call..."

success = tapiline.DropCallAsynch
If success <> True Then
    MsgBox "Error #" & tapiline.LastError & vbCrLf & tapiline.ErrorString(tapiline.LastError)
End If
    

End Sub




'THESE ARE EVENTS RAISED BY THE CLASS
'WHEN ASYNCH FUNCTIONS RETURN
Private Sub tapiline_DropCallResult(ByVal errorCode As Long)
    If errorCode = TAPI_SUCCESS Then
        lstCallProgress.AddItem "Call Dropped: " & txtPhoneNumber.Text
        lstCallProgress.TopIndex = lstCallProgress.ListCount - 1
    Else
        lstCallProgress.AddItem "Call Drop Error: " & tapiline.ErrorString(errorCode)
        lstCallProgress.TopIndex = lstCallProgress.ListCount - 1
    End If
    
End Sub


Private Sub tapiline_MakeCallResult(ByVal errorCode As Long)
    If errorCode = TAPI_SUCCESS Then
        lstCallProgress.AddItem "Dialing: " & txtPhoneNumber.Text
        lstCallProgress.TopIndex = lstCallProgress.ListCount - 1
    Else
        lstCallProgress.AddItem "Could not dial: " & txtPhoneNumber.Text
        lstCallProgress.AddItem "Dial Error: " & tapiline.ErrorString(errorCode)
        lstCallProgress.TopIndex = lstCallProgress.ListCount - 1
    End If
End Sub
'THESE ARE EVENTS RAISED BY THE CLASS WHEN IT RECEIVES
'CERTAIN LINE_STATE MSGS
Private Sub tapiline_Connected()
cmdHangUp.Enabled = True
lstCallProgress.AddItem "Connected"
lstCallProgress.TopIndex = lstCallProgress.ListCount - 1
End Sub

Private Sub tapiline_Disconnected()
cmdHangUp.Enabled = False
lstCallProgress.AddItem "Disconnected"
lstCallProgress.TopIndex = lstCallProgress.ListCount - 1
End Sub
Private Sub tapiline_Idle()
cmdHangUp.Enabled = False
cmdDial.Enabled = True
lstCallProgress.AddItem "Idle"
lstCallProgress.TopIndex = lstCallProgress.ListCount - 1
End Sub

'this sub is only used to slow down the initial updating of the
'listbox control - it is only for looks
Sub Pause(ByVal mSecs As Long, Optional bYield As Boolean = True)
    Dim startTime As Long
    
    startTime = timeGetTime()
    Do While timeGetTime < startTime + mSecs
        If bYield Then
            DoEvents
        End If
    Loop
    
End Sub
