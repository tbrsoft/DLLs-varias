VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WaveIn Recorder"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   362
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmEffects 
      Caption         =   "Effects/Filters"
      Height          =   990
      Left            =   233
      TabIndex        =   26
      Top             =   4950
      Width           =   4965
      Begin VB.CheckBox chkFX 
         Caption         =   "Equalizer"
         Height          =   240
         Index           =   3
         Left            =   1650
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkFX 
         Caption         =   "Amplifier"
         Height          =   240
         Index           =   2
         Left            =   1650
         TabIndex        =   30
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox chkFX 
         Caption         =   "Phase Shift"
         Height          =   240
         Index           =   1
         Left            =   300
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdFXOpts 
         Caption         =   "Settings"
         Height          =   315
         Left            =   3525
         TabIndex        =   28
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox chkFX 
         Caption         =   "Echo"
         Height          =   240
         Index           =   0
         Left            =   300
         TabIndex        =   27
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4350
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picVUR 
      Height          =   195
      Left            =   1470
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   242
      TabIndex        =   25
      Top             =   4575
      Width           =   3690
   End
   Begin VB.PictureBox picVUL 
      Height          =   195
      Left            =   1470
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   242
      TabIndex        =   23
      Top             =   4350
      Width           =   3690
   End
   Begin ComctlLib.StatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   21
      Top             =   6105
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2893
            MinWidth        =   2893
            Text            =   "Filesize: 0 Bytes"
            TextSave        =   "Filesize: 0 Bytes"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2893
            MinWidth        =   2893
            Text            =   "Length: 0:00"
            TextSave        =   "Length: 0:00"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.Slider sldVol 
      Height          =   345
      Left            =   1245
      TabIndex        =   20
      Top             =   1250
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   609
      _Version        =   327682
      LargeChange     =   2000
      Max             =   65535
      TickStyle       =   3
   End
   Begin VB.ComboBox cboRecLine 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   845
      Width           =   3615
   End
   Begin VB.Timer tmrVis 
      Interval        =   25
      Left            =   4950
      Top             =   0
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "Record"
      Height          =   315
      Left            =   3683
      TabIndex        =   14
      Top             =   3825
      Width           =   1515
   End
   Begin VB.PictureBox picFreq 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   325
      Left            =   233
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   13
      Top             =   3840
      Width           =   1545
   End
   Begin VB.Frame frmInp 
      Caption         =   "Input"
      Height          =   1665
      Left            =   120
      TabIndex        =   7
      Top             =   150
      Width           =   5190
      Begin VB.ComboBox cboDev 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   3615
      End
      Begin VB.Label lblDev 
         Caption         =   "Device:"
         Height          =   240
         Left            =   375
         TabIndex        =   12
         Top             =   345
         Width           =   765
      End
      Begin VB.Label lblLine 
         Caption         =   "Line:"
         Height          =   240
         Left            =   375
         TabIndex        =   11
         Top             =   735
         Width           =   615
      End
      Begin VB.Label lblVol 
         Caption         =   "Volume:"
         Height          =   240
         Left            =   375
         TabIndex        =   10
         Top             =   1125
         Width           =   690
      End
      Begin VB.Label lblVolPer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Left            =   4245
         TabIndex        =   9
         Top             =   1125
         Width           =   435
      End
   End
   Begin VB.Frame frmOut 
      Caption         =   "Output"
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   1950
      Width           =   5190
      Begin VB.PictureBox picOutXP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   75
         ScaleHeight     =   1365
         ScaleWidth      =   5040
         TabIndex        =   2
         Top             =   225
         Width           =   5040
         Begin VB.CommandButton cmdSelFmt 
            Caption         =   "..."
            Height          =   285
            Left            =   4500
            TabIndex        =   18
            Top             =   900
            Width           =   465
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Left            =   4500
            TabIndex        =   17
            Top             =   75
            Width           =   465
         End
         Begin VB.TextBox txtWAVFmt 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   900
            Width           =   3240
         End
         Begin VB.TextBox txtFile 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   75
            Width           =   4290
         End
         Begin VB.ComboBox cboSamplerate 
            Height          =   315
            ItemData        =   "Form1.frx":0000
            Left            =   1200
            List            =   "Form1.frx":0010
            TabIndex        =   4
            Text            =   "cboSamplerate"
            Top             =   450
            Width           =   1290
         End
         Begin VB.CheckBox chkStereo 
            Caption         =   "stereo"
            Height          =   240
            Left            =   2625
            TabIndex        =   3
            Top             =   495
            Width           =   1365
         End
         Begin VB.Label lblWAVFmt 
            AutoSize        =   -1  'True
            Caption         =   "WAV Format:"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   930
            Width           =   960
         End
         Begin VB.Label lblSamplerate 
            AutoSize        =   -1  'True
            Caption         =   "Samplerate:"
            Height          =   195
            Left            =   150
            TabIndex        =   6
            Top             =   510
            Width           =   870
         End
      End
   End
   Begin VB.PictureBox picAmpl 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   325
      Left            =   1883
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   0
      Top             =   3825
      Width           =   1545
   End
   Begin VB.Label lblVUchR 
      AutoSize        =   -1  'True
      Caption         =   "Right Channel:"
      Height          =   195
      Left            =   270
      TabIndex        =   24
      Top             =   4575
      Width           =   1065
   End
   Begin VB.Label lblVUchL 
      AutoSize        =   -1  'True
      Caption         =   "Left Channel:"
      Height          =   195
      Left            =   270
      TabIndex        =   22
      Top             =   4350
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' VU gradient colors
Private Const COLOR_START       As Long = vbGreen
Private Const COLOR_MIDDLE      As Long = &H22A8E1
Private Const COLOR_END         As Long = vbRed

' the bigger, the smoother
Private Const VU_SMOOTH_LEN     As Long = 3

Private WithEvents clsRecorder  As WaveInRecorder
Attribute clsRecorder.VB_VarHelpID = -1

Private clsVUSmoothL            As clsSmooth
Private clsVUSmoothR            As clsSmooth
Private clsVis                  As clsDraw
Private clsDSP                  As clsDSP
Private intSamples()            As Integer

Private clsEncoder              As EncoderWAV

Private lngMSEncoded            As Long
Private lngBytesPerSec          As Long

Private blnLoaded               As Boolean

Private Sub cboDev_Click()
    cboRecLine.Clear

    If Not clsRecorder.SelectDevice(cboDev.ListIndex) Then
        MsgBox "Couldn't select device!", vbExclamation
        Exit Sub
    End If

    ShowLines
End Sub

Private Sub cboRecLine_Click()
    If Not clsRecorder.SelectMixerLine(cboRecLine.ListIndex) Then
        MsgBox "Couldn't select mixer line!", vbExclamation
    End If

    ' MixerLineType can be used to automaticaly find and set
    ' the line you want to record from, e.g. microphone.
    ' MixerLine also accepts a line id as a parameter,
    ' pass -1 and the currently selected line is returned.
    Debug.Print "Line Type: ";
    Select Case clsRecorder.MixerLineType
        Case MIXERLINE_ANALOG:      Debug.Print "Analog"
        Case MIXERLINE_AUXILIARY:   Debug.Print "Auxiliary"
        Case MIXERLINE_COMPACTDISC: Debug.Print "Compact Disc"
        Case MIXERLINE_DIGITAL:     Debug.Print "Digital"
        Case MIXERLINE_LINE:        Debug.Print "Line-In"
        Case MIXERLINE_MICROPHONE:  Debug.Print "Microphone"
        Case MIXERLINE_PCSPEAKER:   Debug.Print "PC Speaker"
        Case MIXERLINE_SYNTHESIZER: Debug.Print "Synthesizer"
        Case MIXERLINE_TELEPHONE:   Debug.Print "Telephone"
        Case MIXERLINE_UNDEFINED:   Debug.Print "Undefined"
        Case MIXERLINE_WAVEOUT:     Debug.Print "WaveOut"
        Case Else:                  Debug.Print "Unknown"
    End Select

    sldVol.value = clsRecorder.MixerLineVolume
    sldVol_Scroll
End Sub

Private Sub cboSamplerate_Click()
    ' the currently selected WAV output codec
    ' couldn't support resampling, so force the
    ' user to select a new format which is
    ' compatible to the new samplerate
    If blnLoaded Then
        cmdSelFmt_Click
        clsDSP.samplerate = CLng(cboSamplerate.Text)
    End If
End Sub

Private Sub cboSamplerate_KeyPress( _
    KeyAscii As Integer _
)

    If KeyAscii <> vbKeyBack Then
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        Else
            If Len(cboSamplerate.Text) = 5 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub chkFX_Click( _
    Index As Integer _
)

    If chkFX(Index).value Then
        clsDSP.EffectsUsed = clsDSP.EffectsUsed Or (2 ^ Index)
    Else
        clsDSP.EffectsUsed = clsDSP.EffectsUsed And (Not (2 ^ Index))
    End If
End Sub

Private Sub chkStereo_Click()
    ' the currently selected WAV output codec
    ' couldn't support channel mixing, so force the
    ' user to select a new format which is
    ' compatible to the new channels count
    If blnLoaded Then
        cmdSelFmt_Click
        clsDSP.Channels = chkStereo.value + 1
    End If
End Sub

' waveIn returns a buffer to the app
Private Sub clsRecorder_GotData( _
    intBuffer() As Integer, _
    lngLen As Long _
)

    ' save the current buffer for visualizing it
    intSamples = intBuffer

    clsDSP.ProcessSamples intSamples

    ' GotData could also be raised after recording
    ' got stopped because a buffer was just finished
    ' when StopRecord got called
    If Not clsRecorder.IsRecording Then Exit Sub

    ' update recorded time
    lngMSEncoded = lngMSEncoded + ((lngLen / lngBytesPerSec) * 1000)

    If Not clsEncoder Is Nothing Then
        ' send PCM data to the WAV encoder
        If clsEncoder.Encoder_Encode(VarPtr(intSamples(0)), lngLen, 0) = SND_ERR_WRITE_ERROR Then
            cmdDo_Click
            MsgBox "Write error (no disk space left?)!", vbExclamation
        End If
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim strExt  As String

    strExt = "*." & clsEncoder.Encoder_Extension

    With dlg
        .FileName = vbNullString
        .Flags = cdlOFNOverwritePrompt
        .Filter = clsEncoder.Encoder_Description & " (" & strExt & ")|" & strExt
        .ShowOpen
    End With

    If dlg.FileName <> vbNullString Then
        If Not Right$(LCase$(dlg.FileName), 4) = ".wav" Then
            txtFile.Text = dlg.FileName & ".wav"
        Else
            txtFile.Text = dlg.FileName
        End If
    End If
End Sub

Private Sub cmdDo_Click()
    Dim sndres  As SND_RESULT

    If clsRecorder.IsRecording Then
        If Not clsRecorder.StopRecord Then
            MsgBox "Could not stop recording!", vbExclamation
        End If

        clsEncoder.Encoder_EncoderClose

        lngMSEncoded = 0
        lngBytesPerSec = 0

        cmdDo.Caption = "Start recording"
        SetFormEnabled True
    Else
        If txtFile.Text = "" Then
            MsgBox "Output file missing!", vbExclamation
            Exit Sub
        End If

        If cboDev.ListIndex = -1 Then
            MsgBox "No device selected!", vbExclamation
            Exit Sub
        End If

        ' Blockalign = (Bits/Sample / 8) * Channels
        ' Bytes/Sec = Samplerate * Blockalign
        '
        ' we simply use 16 bit samples (integers) where we can,
        ' that's the most comfortable way to work with
        ' samples in VB
        lngBytesPerSec = (CLng(cboSamplerate.Text) * (2 * (chkStereo.value + 1)))

        sndres = clsEncoder.Encoder_EncoderInit(CLng(cboSamplerate.Text), _
                                                chkStereo.value + 1, _
                                                txtFile.Text)

        If sndres <> SND_ERR_SUCCESS Then
            MsgBox "Could not init the encoder!", vbExclamation
            Exit Sub
        End If

        clsDSP.samplerate = CLng(cboSamplerate.Text)
        clsDSP.Channels = chkStereo.value + 1

        If Not clsRecorder.StartRecord(cboSamplerate.Text, chkStereo.value + 1) Then
            MsgBox "Could not start recording!", vbExclamation
        End If

        cmdDo.Caption = "Stop recording"
        SetFormEnabled False
    End If
End Sub

Private Sub SetFormEnabled( _
    bln As Boolean _
)

    frmOut.Enabled = bln
    frmInp.Enabled = bln
End Sub

Private Sub cmdFXOpts_Click()
    frmFXOpts.ShowEx clsDSP, Me
End Sub

Private Sub cmdSelFmt_Click()
    clsEncoder.SelectFormat CLng(cboSamplerate.Text), _
                            chkStereo.value + 1, _
                            Me.hWnd, _
                            "Select WAV output format"

    UpdateWAVFmtDisp
End Sub

Private Sub UpdateWAVFmtDisp()
    With clsEncoder
        txtWAVFmt.Text = .FormatTag & " - " & .FormatID
    End With
End Sub

Private Sub Form_Load()
    Set clsVUSmoothL = New clsSmooth
    Set clsVUSmoothR = New clsSmooth
    Set clsRecorder = New WaveInRecorder
    Set clsEncoder = New EncoderWAV
    Set clsVis = New clsDraw
    Set clsDSP = New clsDSP

    ' buffer 3 peaks for the VU meter
    clsVUSmoothL.SmoothNew VU_SMOOTH_LEN
    clsVUSmoothR.SmoothNew VU_SMOOTH_LEN

    ' std. format is 44.1 kHz stereo (16 bit, of course)
    cboSamplerate.Text = "44100"
    chkStereo.value = 1

    ReDim intSamples(FFT_SAMPLES - 1) As Integer

    blnLoaded = True

    UpdateWAVFmtDisp
    ShowDevices

    clsDSP.samplerate = CLng(cboSamplerate.Text)
    clsDSP.Channels = chkStereo.value + 1

    Me.Show
End Sub

Private Sub ShowDevices()
    Dim i   As Long

    cboDev.Clear

    For i = 0 To clsRecorder.DeviceCount - 1
        cboDev.AddItem clsRecorder.DeviceName(i)
    Next
End Sub

Private Sub ShowLines()
    Dim i   As Long

    cboRecLine.Clear

    For i = 0 To clsRecorder.MixerLineCount - 1
        cboRecLine.AddItem clsRecorder.MixerLineName(i)
    Next

    cboRecLine.ListIndex = clsRecorder.SelectedMixerLine
End Sub

Private Sub Form_Unload( _
    Cancel As Integer _
)

    If clsRecorder.IsRecording Then
        cmdDo_Click
    End If

    Set clsRecorder = Nothing
    Set clsEncoder = Nothing
End Sub

Private Sub sldVol_Click()
    sldVol_Scroll
End Sub

Private Sub sldVol_Scroll()
    clsRecorder.MixerLineVolume = sldVol.value
    lblVolPer.Caption = Fix(sldVol.value / sldVol.max * 100) & "%"
End Sub

Private Sub tmrVis_Timer()
    Dim lngMaxL As Long
    Dim lngMaxR As Long

    If clsRecorder.IsRecording Then
        ' frequency spectrum
        clsVis.DrawAmplitudes intSamples, picAmpl
        ' amplitude curve
        clsVis.DrawFrequencies intSamples, picFreq

        ' VU meter
        If chkStereo.value = 1 Then
            lngMaxL = GetArrayMaxAbs(intSamples, 0, 2)
            lngMaxR = GetArrayMaxAbs(intSamples, 1, 2)
        Else
            lngMaxL = GetArrayMaxAbs(intSamples)
            lngMaxR = lngMaxL
        End If

        If lngMaxL = 0 Then lngMaxL = 1
        If lngMaxR = 0 Then lngMaxR = 1

        clsVUSmoothL.SmoothAdd lngMaxL
        clsVUSmoothR.SmoothAdd lngMaxR

        DrawVU picVUL, clsVUSmoothL.SmoothGetMax / 32768#
        DrawVU picVUR, clsVUSmoothR.SmoothGetMax / 32768#

        ' Info
        sbar.Panels(1).Text = "Filesize: " & FormatFileSize(FileLen(txtFile.Text))
        sbar.Panels(2).Text = "Length: " & FmtTime(lngMSEncoded)
    Else
        sbar.Panels(1).Text = "Filesize: 0 Bytes"
        sbar.Panels(2).Text = "Length: 0:00"
    End If
End Sub

Private Function FmtTime( _
    ByVal lngMS As Long _
) As String

    Dim lngMin  As Long
    Dim lngSec  As Long

    lngSec = lngMS / 1000
    lngMin = lngSec \ 60
    lngSec = lngSec Mod 60

    FmtTime = lngMin & ":" & Format(lngSec, "00")
End Function

Private Function FormatFileSize( _
    ByVal dblFileSize As Double, _
    Optional ByVal strFormatMask As String _
) As String

    Select Case dblFileSize
        Case 0 To 1023 ' Bytes
            FormatFileSize = Format(dblFileSize) & " bytes"

        Case 1024 To 1048575 ' KB
            If strFormatMask = Empty Then strFormatMask = "###0"
            FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"

        Case 1024# ^ 2 To 1073741823 ' MB
            If strFormatMask = Empty Then strFormatMask = "###0.0"
            FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"

        Case Is > 1073741823# ' GB
            If strFormatMask = Empty Then strFormatMask = "###0.0"
            FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
    End Select
End Function

' get the greatest absolute value in an array of samples
Private Function GetArrayMaxAbs( _
    intArray() As Integer, _
    Optional ByVal offStart As Long = 0, _
    Optional ByVal steps As Long = 1 _
) As Long

    Dim lngTemp As Long
    Dim lngMax  As Long
    Dim i       As Long

    For i = offStart To UBound(intArray) Step steps
        lngTemp = Abs(CLng(intArray(i)))
        If lngTemp > lngMax Then
            lngMax = lngTemp
        End If
    Next

    GetArrayMaxAbs = lngMax
End Function

Private Sub DrawVU( _
    ByVal picbox As PictureBox, _
    ByVal value As Single _
)

    Dim lngColor    As Long

    lngColor = GetGradColor(1, value, COLOR_START, COLOR_MIDDLE, COLOR_END)

    picbox.Cls
    clsVis.DrawRect picbox.hdc, 0, 0, value * picbox.Width, picbox.Height, lngColor
End Sub
