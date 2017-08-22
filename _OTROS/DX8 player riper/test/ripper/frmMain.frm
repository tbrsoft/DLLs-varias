VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "MonoRipper"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   735
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picEncMP3 
      Height          =   1635
      Left            =   5850
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   25
      Top             =   990
      Width           =   4965
      Begin VB.ComboBox cboMP3VBRQ 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   2880
         List            =   "frmMain.frx":001F
         Style           =   2  'Dropdown-Liste
         TabIndex        =   29
         Top             =   615
         Width           =   1635
      End
      Begin VB.ComboBox cboMP3VBRBit 
         Height          =   315
         ItemData        =   "frmMain.frx":0088
         Left            =   3510
         List            =   "frmMain.frx":00B6
         Style           =   2  'Dropdown-Liste
         TabIndex        =   28
         Top             =   180
         Width           =   1005
      End
      Begin VB.CheckBox chkMP3VBR 
         Caption         =   "VBR"
         Height          =   195
         Left            =   3870
         TabIndex        =   27
         Top             =   1080
         Width           =   645
      End
      Begin VB.ComboBox cboMP3Bit 
         Height          =   315
         ItemData        =   "frmMain.frx":00F9
         Left            =   720
         List            =   "frmMain.frx":0127
         Style           =   2  'Dropdown-Liste
         TabIndex        =   26
         Top             =   180
         Width           =   915
      End
      Begin VB.Label lblMP3VBRQ 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "VBR Quality:"
         Height          =   195
         Left            =   1890
         TabIndex        =   32
         Top             =   675
         Width           =   900
      End
      Begin VB.Label lblMP3VBRMaxBit 
         AutoSize        =   -1  'True
         Caption         =   "Max. VBR Bitrate:"
         Height          =   195
         Left            =   2160
         TabIndex        =   31
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label lblMP3Bit 
         AutoSize        =   -1  'True
         Caption         =   "Bitrate:"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   225
         Width           =   495
      End
   End
   Begin VB.PictureBox picEncAPE 
      Height          =   1635
      Left            =   5850
      ScaleHeight     =   1575
      ScaleWidth      =   4905
      TabIndex        =   15
      Top             =   990
      Width           =   4965
      Begin VB.ComboBox cboAPELevel 
         Height          =   315
         ItemData        =   "frmMain.frx":016A
         Left            =   1800
         List            =   "frmMain.frx":017D
         Style           =   2  'Dropdown-Liste
         TabIndex        =   16
         Top             =   360
         Width           =   2625
      End
      Begin VB.Label lblAPECompLevel 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   195
         Left            =   1275
         TabIndex        =   17
         Top             =   390
         Width           =   435
      End
   End
   Begin VB.PictureBox picEncOGG 
      Height          =   1635
      Left            =   5850
      ScaleHeight     =   1575
      ScaleWidth      =   4905
      TabIndex        =   18
      Top             =   990
      Width           =   4965
      Begin VB.OptionButton optOggBit 
         Caption         =   "Bitrate:"
         Height          =   195
         Left            =   270
         TabIndex        =   21
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton optOggQual 
         Caption         =   "Quality:"
         Height          =   195
         Left            =   2340
         TabIndex        =   20
         Top             =   270
         Width           =   1095
      End
      Begin VB.ComboBox cboOggBitNom 
         Height          =   315
         ItemData        =   "frmMain.frx":01AC
         Left            =   1080
         List            =   "frmMain.frx":01D1
         Style           =   2  'Dropdown-Liste
         TabIndex        =   19
         Top             =   585
         Width           =   1005
      End
      Begin MSComctlLib.Slider sldOggQual 
         Height          =   210
         Left            =   2520
         TabIndex        =   22
         Top             =   600
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   370
         _Version        =   393216
      End
      Begin VB.Label lblOggNominal 
         AutoSize        =   -1  'True
         Caption         =   "Nominal:"
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   630
         Width           =   705
      End
      Begin VB.Label lblOggQual 
         Alignment       =   1  'Rechts
         Caption         =   "Quality: 0.0"
         Height          =   195
         Left            =   3420
         TabIndex        =   23
         Top             =   900
         Width           =   945
      End
   End
   Begin VB.PictureBox picEncWMA 
      Height          =   1635
      Left            =   5850
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   10
      Top             =   990
      Width           =   4965
      Begin VB.ComboBox cboWMACodecs 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown-Liste
         TabIndex        =   12
         Top             =   270
         Width           =   3525
      End
      Begin VB.ComboBox cboWMAFormat 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown-Liste
         TabIndex        =   11
         Top             =   675
         Width           =   3525
      End
      Begin VB.Label lblWMACodec 
         AutoSize        =   -1  'True
         Caption         =   "Codec:"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   315
         Width           =   510
      End
      Begin VB.Label lblWMAFormat 
         AutoSize        =   -1  'True
         Caption         =   "Format:"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   720
         Width           =   570
      End
   End
   Begin VB.PictureBox picEncWAV 
      Height          =   1635
      Left            =   5850
      ScaleHeight     =   1575
      ScaleWidth      =   4905
      TabIndex        =   5
      Top             =   990
      Width           =   4965
      Begin VB.CommandButton cmdWAVChangeFmt 
         Caption         =   "change"
         Height          =   285
         Left            =   3330
         TabIndex        =   7
         Top             =   810
         Width           =   1455
      End
      Begin VB.CheckBox chkWAVWriteHdr 
         Caption         =   "Write WAV Header"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   810
         Value           =   1  'Aktiviert
         Width           =   2175
      End
      Begin VB.Label lblWAVFormat 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Format"
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   450
         Width           =   4695
      End
      Begin VB.Label lblWAVACMFormat 
         AutoSize        =   -1  'True
         Caption         =   "ACM Format:"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdBrowsePath 
      Caption         =   "..."
      Height          =   285
      Left            =   10260
      TabIndex        =   37
      Top             =   3150
      Width           =   465
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3150
      Width           =   4425
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4410
      TabIndex        =   4
      Top             =   3960
      Width           =   1185
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Rip"
      Height          =   285
      Left            =   4410
      TabIndex        =   3
      Top             =   3600
      Width           =   1185
   End
   Begin MSComctlLib.ProgressBar prgCurrent 
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   3660
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwTracks 
      Height          =   2805
      Left            =   270
      TabIndex        =   1
      Top             =   630
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   4948
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Track"
         Object.Width           =   2126
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Länge"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Typ"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.ComboBox cboDevices 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   180
      Width           =   5325
   End
   Begin MSComctlLib.TabStrip tabEncoders 
      Height          =   2085
      Left            =   5760
      TabIndex        =   33
      Top             =   630
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   3678
      TabMinWidth     =   1688
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "WAV"
            Key             =   "WAV"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MP3"
            Key             =   "MP3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OGG"
            Key             =   "OGG"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "APE"
            Key             =   "APE"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "WMA"
            Key             =   "WMA"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prgTotal 
      Height          =   195
      Left            =   270
      TabIndex        =   39
      Top             =   4050
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "total:"
      Height          =   195
      Left            =   270
      TabIndex        =   40
      Top             =   3855
      Width           =   345
   End
   Begin VB.Label lblCurrent 
      AutoSize        =   -1  'True
      Caption         =   "current:"
      Height          =   195
      Left            =   270
      TabIndex        =   38
      Top             =   3465
      Width           =   540
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5760
      TabIndex        =   35
      Top             =   2880
      Width           =   525
   End
   Begin VB.Label lblEncoder 
      AutoSize        =   -1  'True
      Caption         =   "Encoder:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5760
      TabIndex        =   34
      Top             =   270
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IDSPCallback

Private Type MSF
    m               As Byte
    s               As Byte
    F               As Byte
End Type

Private clsEncWAV   As Monoton_DS.EncoderWAV
Private clsEncAPE   As Monoton_DS.EncoderAPE
Private clsEncMP3   As Monoton_DS.EncoderMP3
Private clsEncOGG   As Monoton_DS.EncoderOGG
Private clsEncWMA   As Monoton_DS.EncoderWMA

Private clsCDInfo   As Monoton_DS.StreamCDA
Private clsRipper   As Monoton_DS.ISoundStream
Private clsPCM      As Monoton_DS.PCMPreparator
Private clsEnc      As Monoton_DS.IEncoder

Private blnCancel   As Boolean

' *********************************************
' * UI
' *********************************************

Private Sub cmdBrowsePath_Click()
    txtPath.Text = AddSlash(BrowseForFolder(Me.hWnd, "Neuer Pfad"))
End Sub

Private Sub cmdCancel_Click()
    blnCancel = True
End Sub

Private Sub cmdStart_Click()
    Dim i               As Long
    Dim lngSize         As Long
    Dim lngRead         As Long
    Dim btData()        As Byte
    Dim strOutputFile   As String
    Dim strInputFile    As String

    prgTotal.Max = 1
    prgTotal.Value = 0

    For i = 1 To lvwTracks.ListItems.Count
        If lvwTracks.ListItems(i).Checked Then
            prgTotal.Max = prgTotal.Max + 1
        End If
    Next

    If prgTotal.Max = 1 Then
        MsgBox "You need to select at least 1 track!", vbExclamation
        Exit Sub
    End If

    prgTotal.Max = prgTotal.Max - 1

    cmdStart.Enabled = Not cmdStart.Enabled
    cmdCancel.Enabled = Not cmdCancel.Enabled

    Select Case tabEncoders.SelectedItem.Key
        Case "WAV": Set clsEnc = clsEncWAV
        Case "MP3": Set clsEnc = clsEncMP3
        Case "APE": Set clsEnc = clsEncAPE
        Case "WMA": Set clsEnc = clsEncWMA
        Case "OGG": Set clsEnc = clsEncOGG
    End Select

    For i = 1 To lvwTracks.ListItems.Count

        strOutputFile = txtPath.Text & "Track " & Format(i, 0) & "." & clsEnc.Extension
        strInputFile = clsCDInfo.DeviceChar(cboDevices.ListIndex) & ":\" & "Track" & Format(i, "00") & ".cda"

        With lvwTracks.ListItems(i)
            If .Checked Then

                If clsRipper.OpenSource(strInputFile) <> STREAM_OK Then
                    Debug.Print "Could not open: " & strInputFile
                    GoTo SkipItem
                End If

                If Not clsPCM.InitConversion(clsRipper, Me) Then
                    Debug.Print "Could not init PCMPreparator"
                    GoTo SkipItem
                End If

                If clsEnc.Init(strOutputFile, 44100, 2, 16, lngSize) <> STREAM_OK Then
                    Debug.Print "Could not init encoder (Track " & i & ")"
                    GoTo SkipItem
                End If

                ReDim btData(lngSize - 1) As Byte

                Do
                    clsPCM.GetSamples VarPtr(btData(0)), lngSize, lngRead
                    If clsPCM.EndOfStream Then
                        If lngRead > 0 Then
                            clsEnc.Encode VarPtr(btData(0)), lngRead
                        End If
                        Exit Do
                    End If

                    If Not clsEnc.Encode(VarPtr(btData(0)), lngRead) = STREAM_OK Then
                        Exit Do
                    End If

                    If blnCancel Then
                        Exit Do
                    End If

                    prgCurrent.Value = Min(clsRipper.Info.position / clsRipper.Info.Duration * 100, 100)

                    DoEvents
                Loop

SkipItem:
                clsPCM.CloseConverter
                clsRipper.CloseSource
                clsEnc.DeInit

                prgTotal.Value = prgTotal.Value + 1

                If blnCancel Then
                    blnCancel = False
                    Exit For
                End If

            End If
        End With

    Next

    cmdStart.Enabled = Not cmdStart.Enabled
    cmdCancel.Enabled = Not cmdCancel.Enabled

    MsgBox "Finished!", vbInformation
End Sub

Private Sub IDSPCallback_Samples( _
    intSamples() As Integer, _
    ByVal datalength As Long, _
    ByVal channels As Integer _
)
    '
End Sub

Private Sub tabEncoders_Click()
    Select Case tabEncoders.SelectedItem.Key
        Case "WAV"
            picEncWAV.Visible = True
            picEncMP3.Visible = False
            picEncOGG.Visible = False
            picEncAPE.Visible = False
            picEncWMA.Visible = False
        Case "MP3"
            picEncWAV.Visible = False
            picEncMP3.Visible = True
            picEncOGG.Visible = False
            picEncAPE.Visible = False
            picEncWMA.Visible = False
        Case "OGG"
            picEncWAV.Visible = False
            picEncMP3.Visible = False
            picEncOGG.Visible = True
            picEncAPE.Visible = False
            picEncWMA.Visible = False
        Case "APE"
            picEncWAV.Visible = False
            picEncMP3.Visible = False
            picEncOGG.Visible = False
            picEncAPE.Visible = True
            picEncWMA.Visible = False
        Case "WMA"
            picEncWAV.Visible = False
            picEncMP3.Visible = False
            picEncOGG.Visible = False
            picEncAPE.Visible = False
            picEncWMA.Visible = True
    End Select
End Sub

Private Sub cboDevices_Click()
    lvwTracks.ListItems.Clear

    If clsCDInfo.SelectDevice(cboDevices.ListIndex) Then
        ShowTracks
    Else
        MsgBox "Could not select the drive!", vbExclamation
    End If
End Sub

Private Sub Form_Load()
    Set clsPCM = New Monoton_DS.PCMPreparator
    Set clsCDInfo = New Monoton_DS.StreamCDA
    Set clsRipper = clsCDInfo

    ' Encoder
    InitWAVEncoder
    InitMP3Encoder
    InitOGGEncoder
    InitAPEEncoder
    InitWMAEncoder

    tabEncoders_Click
    txtPath.Text = AddSlash(App.Path)

    cboDevices.Clear
    ShowDevices
End Sub

' *********************************************
' * MP3 settings
' *********************************************

Private Sub chkMP3VBR_Click()
    clsEncMP3.VBR = CBool(chkMP3VBR.Value)
End Sub

Private Sub cboMP3Bit_Click()
    clsEncMP3.Bitrate = CLng(cboMP3Bit.List(cboMP3Bit.ListIndex))
End Sub

Private Sub cboMP3VBRBit_Click()
    clsEncMP3.VBRMaxBitrate = CLng(cboMP3VBRBit.List(cboMP3VBRBit.ListIndex))
End Sub

Private Sub cboMP3VBRQ_Click()
    clsEncMP3.VBRQuality = cboMP3VBRQ.ListIndex
End Sub

' *********************************************
' * APE settings
' *********************************************

Private Sub cboAPELevel_Click()
    Select Case cboAPELevel.ListIndex
        Case 0: clsEncAPE.CompressionLevel = COMPRESSION_LEVEL_FAST
        Case 1: clsEncAPE.CompressionLevel = COMPRESSION_LEVEL_NORMAL
        Case 2: clsEncAPE.CompressionLevel = COMPRESSION_LEVEL_HIGH
        Case 3: clsEncAPE.CompressionLevel = COMPRESSION_LEVEL_EXTRA_HIGH
        Case 4: clsEncAPE.CompressionLevel = COMPRESSION_LEVEL_INSANE
    End Select
End Sub

' *********************************************
' * WMA settings
' *********************************************

Private Sub cboWMACodecs_Click()
    Dim i   As Long

    cboWMAFormat.Clear

    For i = 0 To clsEncWMA.CodecFormatCount(cboWMACodecs.ListIndex) - 1
        cboWMAFormat.AddItem clsEncWMA.CodecFormatName(cboWMACodecs.ListIndex, i)
    Next

    If cboWMAFormat.ListCount > 0 Then
        cboWMAFormat.ListIndex = 0
    End If
End Sub

Private Sub cboWMAFormat_Click()
    If Not clsEncWMA.SelectFormat(cboWMACodecs.ListIndex, cboWMAFormat.ListIndex) Then
        MsgBox "Couldn't select format!", vbExclamation
    End If
End Sub

' *********************************************
' * WAV settings
' *********************************************

Private Sub chkWAVWriteHdr_Click()
    clsEncWAV.WriteWAVHeader = CBool(chkWAVWriteHdr.Value)
End Sub

Private Sub cmdWAVChangeFmt_Click()
    If clsEncWAV.SelectFormat(44100, 2, 16, Me.hWnd) = STREAM_OK Then
        lblWAVFormat.Caption = clsEncWAV.FormatTag & " - " & clsEncWAV.FormatID
    End If
End Sub

' *********************************************
' * OGG Vorbis settings
' *********************************************

Private Sub sldOggQual_Change()
    clsEncOGG.Quality = sldOggQual.Value / 10
    lblOggQual.Caption = "Quality: " & clsEncOGG.Quality
End Sub

Private Sub optOggBit_Click()
    clsEncOGG.EncoderMode = OV_ENC_ABR
End Sub

Private Sub optOggQual_Click()
    clsEncOGG.EncoderMode = OV_ENC_QUALITY
End Sub

' *********************************************
' * Encoder settings
' *********************************************

Private Sub InitWAVEncoder()
    Set clsEncWAV = New Monoton_DS.EncoderWAV

    With clsEncWAV
        lblWAVFormat.Caption = .FormatTag & " - " & .FormatID
        chkWAVWriteHdr.Value = Abs(.WriteWAVHeader)
    End With
End Sub

Private Sub InitMP3Encoder()
    Set clsEncMP3 = New Monoton_DS.EncoderMP3

    cboMP3Bit.ListIndex = 0
    cboMP3VBRBit.ListIndex = 0
    cboMP3VBRQ.ListIndex = 0
    chkMP3VBR.Value = 0
End Sub

Private Sub InitOGGEncoder()
    Set clsEncOGG = New Monoton_DS.EncoderOGG

    If clsEncOGG.EncoderMode = OV_ENC_ABR Then
        optOggBit.Value = True
    Else
        optOggQual.Value = True
    End If

    cboOggBitNom.Text = clsEncOGG.BitrateNominal / 1000

    sldOggQual.Value = clsEncOGG.Quality * 10
    lblOggQual = "Quality: " & clsEncOGG.Quality
End Sub

Private Sub InitAPEEncoder()
    Set clsEncAPE = New Monoton_DS.EncoderAPE

    cboAPELevel.ListIndex = 0
End Sub

Private Sub InitWMAEncoder()
    Dim i   As Long

    Set clsEncWMA = New Monoton_DS.EncoderWMA

    For i = 0 To clsEncWMA.CodecsCount - 1
        cboWMACodecs.AddItem clsEncWMA.codecname(i)
    Next
    cboWMACodecs.ListIndex = 0
End Sub

' *********************************************
' * helpers
' *********************************************

Private Sub ShowDevices()
    Dim i   As Long

    With clsCDInfo
        For i = 0 To .DeviceCount - 1
            cboDevices.AddItem .DeviceChar(i) & ": " & .DeviceName(i)
        Next
    End With

    cboDevices.ListIndex = 0
End Sub

Private Sub ShowTracks()
    On Error GoTo ErrorHandler

    Dim i       As Long
    Dim lstitm  As ListItem

    With clsCDInfo.toc
        For i = 1 To .TrackCount - 1
            Set lstitm = lvwTracks.ListItems.Add(Text:="Track " & Format(.track(i).TrackNumber, "00"))
            lstitm.SubItems(1) = MSF2STR(LBA2MSF(.track(i + 1).StartLBA - .track(i).StartLBA))
            lstitm.SubItems(2) = IIf(.track(i).IsAudio, "Audio", "Daten")
            lstitm.Checked = .track(i).IsAudio
        Next
    End With

ErrorHandler:
End Sub

' sectors to Minutes:Seconds:Frames
Private Function LBA2MSF( _
    ByVal LBA As Long _
) As MSF

    Dim m As Long, s As Long, F As Long, Start As Long

    Start = Choose(Abs(CBool(LBA >= -150)) + 1, 450150, 150)

    With LBA2MSF
        .m = Fix((LBA + Start) / (60& * 75&))
        .s = Fix((LBA + Start - .m * 60& * 75&) / 75&)
        .F = Fix(LBA + Start - .m * 60& * 75& - .s * 75&)
    End With
End Function

Private Function MSF2STR( _
    fmt As MSF _
) As String

    MSF2STR = Format(fmt.m, "00") & ":" & _
              Format(fmt.s, "00") & ":" & _
              Format(fmt.F, "00")
End Function

Private Function AddSlash(ByVal strText As String) As String
    AddSlash = IIf(Right$(strText, 1) = "\", strText, strText & "\")
End Function

Private Function Min(val1 As Long, val2 As Long) As Long
    Min = IIf(val1 < val2, val1, val2)
End Function
