VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Monoton Format Transcoder"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
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
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picEncWAV 
      Height          =   1635
      Left            =   323
      ScaleHeight     =   1575
      ScaleWidth      =   4905
      TabIndex        =   4
      Top             =   2430
      Width           =   4965
      Begin VB.CheckBox chkWAVWriteHdr 
         Caption         =   "Write WAV Header"
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   810
         Value           =   1  'Aktiviert
         Width           =   2175
      End
      Begin VB.CommandButton cmdWAVChangeFmt 
         Caption         =   "change"
         Height          =   285
         Left            =   3330
         TabIndex        =   6
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label lblWAVACMFormat 
         AutoSize        =   -1  'True
         Caption         =   "ACM Format:"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   180
         Width           =   945
      End
      Begin VB.Label lblWAVFormat 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Format"
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   450
         Width           =   4695
      End
   End
   Begin VB.PictureBox picEncOGG 
      Height          =   1635
      Left            =   323
      ScaleHeight     =   1575
      ScaleWidth      =   4905
      TabIndex        =   19
      Top             =   2430
      Width           =   4965
      Begin VB.ComboBox cboOggBitNom 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   1080
         List            =   "frmMain.frx":0025
         Style           =   2  'Dropdown-Liste
         TabIndex        =   22
         Top             =   585
         Width           =   1005
      End
      Begin VB.OptionButton optOggQual 
         Caption         =   "Quality:"
         Height          =   195
         Left            =   2340
         TabIndex        =   21
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton optOggBit 
         Caption         =   "Bitrate:"
         Height          =   195
         Left            =   270
         TabIndex        =   20
         Top             =   270
         Width           =   1095
      End
      Begin MSComctlLib.Slider sldOggQual 
         Height          =   210
         Left            =   2520
         TabIndex        =   23
         Top             =   600
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   370
         _Version        =   393216
      End
      Begin VB.Label lblOggQual 
         Alignment       =   1  'Rechts
         Caption         =   "Quality: 0.0"
         Height          =   195
         Left            =   3420
         TabIndex        =   25
         Top             =   900
         Width           =   945
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
   End
   Begin VB.PictureBox picEncMP3 
      Height          =   1635
      Left            =   323
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   12
      Top             =   2430
      Width           =   4965
      Begin VB.ComboBox cboMP3VBRQ 
         Height          =   315
         ItemData        =   "frmMain.frx":005A
         Left            =   2880
         List            =   "frmMain.frx":0079
         Style           =   2  'Dropdown-Liste
         TabIndex        =   41
         Top             =   630
         Width           =   1635
      End
      Begin VB.ComboBox cboMP3Bit 
         Height          =   315
         ItemData        =   "frmMain.frx":00E2
         Left            =   720
         List            =   "frmMain.frx":0110
         Style           =   2  'Dropdown-Liste
         TabIndex        =   15
         Top             =   180
         Width           =   915
      End
      Begin VB.CheckBox chkMP3VBR 
         Caption         =   "VBR"
         Height          =   195
         Left            =   3870
         TabIndex        =   14
         Top             =   1080
         Width           =   645
      End
      Begin VB.ComboBox cboMP3VBRBit 
         Height          =   315
         ItemData        =   "frmMain.frx":0153
         Left            =   3510
         List            =   "frmMain.frx":0181
         Style           =   2  'Dropdown-Liste
         TabIndex        =   13
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label lblMP3Bit 
         AutoSize        =   -1  'True
         Caption         =   "Bitrate:"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   225
         Width           =   495
      End
      Begin VB.Label lblMP3VBRMaxBit 
         AutoSize        =   -1  'True
         Caption         =   "Max. VBR Bitrate:"
         Height          =   195
         Left            =   2160
         TabIndex        =   17
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label lblMP3VBRQ 
         AutoSize        =   -1  'True
         Caption         =   "VBR Quality:"
         Height          =   195
         Left            =   1890
         TabIndex        =   16
         Top             =   675
         Width           =   900
      End
   End
   Begin VB.PictureBox picEncAPE 
      Height          =   1635
      Left            =   323
      ScaleHeight     =   1575
      ScaleWidth      =   4905
      TabIndex        =   26
      Top             =   2430
      Width           =   4965
      Begin VB.ComboBox cboAPELevel 
         Height          =   315
         ItemData        =   "frmMain.frx":01C4
         Left            =   1800
         List            =   "frmMain.frx":01D7
         Style           =   2  'Dropdown-Liste
         TabIndex        =   27
         Top             =   360
         Width           =   2625
      End
      Begin VB.Label lblAPECompLevel 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Compression Level:"
         Height          =   195
         Left            =   315
         TabIndex        =   28
         Top             =   390
         Width           =   1395
      End
   End
   Begin VB.PictureBox picEncWMA 
      Height          =   1635
      Left            =   323
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   29
      Top             =   2430
      Width           =   4965
      Begin VB.ComboBox cboWMAFormat 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown-Liste
         TabIndex        =   33
         Top             =   675
         Width           =   3525
      End
      Begin VB.ComboBox cboWMACodecs 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown-Liste
         TabIndex        =   31
         Top             =   270
         Width           =   3525
      End
      Begin VB.Label lblWMAFormat 
         AutoSize        =   -1  'True
         Caption         =   "Format:"
         Height          =   195
         Left            =   300
         TabIndex        =   32
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblWMACodec 
         AutoSize        =   -1  'True
         Caption         =   "Codec:"
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   315
         Width           =   510
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   990
      TabIndex        =   40
      Top             =   4590
      Width           =   1635
   End
   Begin VB.TextBox txtAlbum 
      Height          =   285
      Left            =   3600
      TabIndex        =   38
      Top             =   4230
      Width           =   1635
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   990
      TabIndex        =   36
      Top             =   4230
      Width           =   1635
   End
   Begin VB.ListBox lstInfo 
      Height          =   1230
      Left            =   450
      TabIndex        =   34
      Top             =   630
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   5040
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3743
      TabIndex        =   11
      Top             =   5400
      Width           =   1635
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   285
      Left            =   3743
      TabIndex        =   10
      Top             =   5040
      Width           =   1635
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   4913
      TabIndex        =   2
      Top             =   180
      Width           =   465
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   863
      TabIndex        =   1
      Top             =   180
      Width           =   3975
   End
   Begin MSComctlLib.TabStrip tabEncoders 
      Height          =   2085
      Left            =   240
      TabIndex        =   3
      Top             =   2070
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
   Begin VB.Label Label1 
      Caption         =   "Title:"
      Height          =   195
      Left            =   360
      TabIndex        =   39
      Top             =   4635
      Width           =   555
   End
   Begin VB.Label lblAlbum 
      Caption         =   "Album:"
      Height          =   195
      Left            =   2880
      TabIndex        =   37
      Top             =   4275
      Width           =   555
   End
   Begin VB.Label lblArtist 
      Caption         =   "Artist:"
      Height          =   195
      Left            =   360
      TabIndex        =   35
      Top             =   4275
      Width           =   555
   End
   Begin VB.Label lblInput 
      Caption         =   "Input:"
      Height          =   285
      Left            =   233
      TabIndex        =   0
      Top             =   180
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IDSPCallback

Private clsEncWAV       As Monoton_DS.EncoderWAV
Private clsEncAPE       As Monoton_DS.EncoderAPE
Private clsEncMP3       As Monoton_DS.EncoderMP3
Private clsEncOGG       As Monoton_DS.EncoderOGG
Private clsEncWMA       As Monoton_DS.EncoderWMA

Private clsStreams()    As Monoton_DS.ISoundStream
Private clsStream       As Monoton_DS.ISoundStream

Private lngStreamCnt    As Long
Private blnCancel       As Boolean

' *********************************************
' * Conversion
' *********************************************

Private Sub cmdCancel_Click()
    blnCancel = True
End Sub

Private Sub cmdConvert_Click()
    Dim clsEncoder      As Monoton_DS.IEncoder
    Dim clsBridge       As Monoton_DS.PCMPreparator
    Dim strFile         As String
    Dim lngReadSize     As Long
    Dim lngRead         As Long
    Dim btData()        As Byte

    Set clsBridge = New Monoton_DS.PCMPreparator

    If txtInput.Text = "" Then
        MsgBox "Keine Datei angegeben!", vbExclamation
        Exit Sub
    End If

    Set clsEncoder = Nothing

    Select Case tabEncoders.SelectedItem.Key
        Case "WAV": Set clsEncoder = clsEncWAV
        Case "MP3": Set clsEncoder = clsEncMP3
        Case "APE": Set clsEncoder = clsEncAPE
        Case "WMA": Set clsEncoder = clsEncWMA
        Case "OGG": Set clsEncoder = clsEncOGG
    End Select

    clsEncoder.Artist = txtArtist.Text
    clsEncoder.Album = txtAlbum.Text
    clsEncoder.Title = txtTitle.Text

    If Not clsBridge.InitConversion(clsStream, Me) Then
        MsgBox "Konnte Konvertierung nicht starten", vbExclamation
        Exit Sub
    End If

    strFile = txtInput.Text & "." & clsEncoder.Extension

    With clsBridge
        If clsEncoder.Init(strFile, .OutputSamplerate, .OutputChannels, .OutputBitsPerSample, lngReadSize) <> STREAM_OK Then
            MsgBox "Konnte Konvertierung nicht starten!", vbExclamation
            Exit Sub
        End If
    End With

    ReDim btData(lngReadSize - 1) As Byte

    cmdConvert.Enabled = Not cmdConvert.Enabled
    cmdCancel.Enabled = Not cmdCancel.Enabled

    Do
        clsBridge.GetSamples VarPtr(btData(0)), lngReadSize, lngRead

        If clsBridge.EndOfStream Then
            If lngRead > 0 Then
                clsEncoder.Encode VarPtr(btData(0)), lngRead
            End If
            Exit Do
        End If

        If Not clsEncoder.Encode(VarPtr(btData(0)), lngRead) = STREAM_OK Then
            Exit Do
        End If

        If blnCancel Then
            blnCancel = False
            Exit Do
        End If

        prg.Value = Min(clsStream.Info.position / clsStream.Info.Duration * 100, 100)

        DoEvents
    Loop

    clsBridge.CloseConverter
    clsEncoder.DeInit

    Set clsEncoder = Nothing
    Set clsBridge = Nothing

    clsStream.SeekTo 0, SEEK_PERCENT

    cmdConvert.Enabled = Not cmdConvert.Enabled
    cmdCancel.Enabled = Not cmdCancel.Enabled

    MsgBox "Fertig!", vbInformation
End Sub

' *********************************************
' * Stream Collection
' *********************************************

Private Function StreamFromExt( _
    ByVal ext As String _
) As ISoundStream

    Dim i       As Long

    ext = Right$(ext, 3)

    For i = 0 To lngStreamCnt - 1
        If InStr(1, Join(clsStreams(i).Extensions, ";"), ext, vbTextCompare) Then
            Set StreamFromExt = clsStreams(i)
            Exit Function
        End If
    Next
End Function

Private Function GetAllExtensions() As String()
    Dim i               As Long
    Dim j               As Long
    Dim strExt          As String
    Dim strExts()       As String
    Dim strCurrExts()   As String

    For i = 0 To lngStreamCnt - 1
        strCurrExts = clsStreams(i).Extensions

        For j = 0 To UBound(strCurrExts)
            If InStr(strExt, strCurrExts(j)) <= 0 Then
                strExt = strExt & strCurrExts(j) & ";"
            End If
        Next
    Next

    If Right$(strExt, 1) = ";" Then
        strExt = Left$(strExt, Len(strExt) - 1)
    End If

    strExts = Split(strExt, ";")

    GetAllExtensions = strExts
End Function

Private Sub AddStream(stream As ISoundStream)
    ReDim Preserve clsStreams(lngStreamCnt) As ISoundStream
    Set clsStreams(lngStreamCnt) = stream
    lngStreamCnt = lngStreamCnt + 1
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
    If clsStream Is Nothing Then
        MsgBox "Please first choose a file!", vbExclamation
        Exit Sub
    End If

    With clsStream.Info
        If clsEncWAV.SelectFormat(.samplerate, .channels, .bitspersample, Me.hWnd) = STREAM_OK Then
            lblWAVFormat.Caption = clsEncWAV.FormatTag & " - " & clsEncWAV.FormatID
        End If
    End With
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
' * others
' *********************************************

Private Sub Form_Load()
    ' Encoder
    InitWAVEncoder
    InitMP3Encoder
    InitOGGEncoder
    InitAPEEncoder
    InitWMAEncoder

    ' Decoder
    AddStream New Monoton_DS.StreamAPE
    AddStream New Monoton_DS.StreamCDA
    AddStream New Monoton_DS.StreamMP3
    AddStream New Monoton_DS.StreamOGG
    AddStream New Monoton_DS.StreamWAV
    AddStream New Monoton_DS.StreamWMA
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

Private Sub cmdBrowse_Click()
    dlg.FileName = vbNullString
    dlg.Filter = Join(GetAllExtensions, "; ") & "|*." & Join(GetAllExtensions, ";*.")
    dlg.ShowOpen
    If dlg.FileName = vbNullString Then Exit Sub

    If Not OpenFile(dlg.FileName) Then
        MsgBox "Could not open file!", vbExclamation
    End If
End Sub

Private Function OpenFile( _
    ByVal strFile As String _
) As Boolean

    Dim i           As Long
    Dim clsTag      As StreamTag
    Dim strArtist   As String
    Dim strAlbum    As String
    Dim strTitle    As String
    Dim strTitleAlt As String

    If Not clsStream Is Nothing Then
        clsStream.CloseSource
    End If

    Set clsStream = Nothing

    Set clsStream = StreamFromExt(strFile)
    If clsStream Is Nothing Then
        Debug.Print "Format not supported"
        Exit Function
    End If

    If clsStream.OpenSource(strFile) <> STREAM_OK Then
        Debug.Print "Could not open file"
        Exit Function
    End If

    txtInput.Text = strFile

    txtAlbum.Text = ""
    txtArtist.Text = ""
    txtTitle.Text = ""

    ' different formats have different tag names
    Select Case UCase$(Right$(strFile, 3))
        Case "WAV", "CDA"
            '
        Case "MP3", "APE", "OGG"
            strArtist = "artist"
            strAlbum = "album"
            strTitle = "title"
            strTitleAlt = ""
        Case "WMA"
            strArtist = "wm/albumartist"
            strAlbum = "wm/albumtitle"
            strTitle = "wm/title"
            strTitleAlt = "title"
    End Select

    With clsStream.Info
        lstInfo.Clear
        lstInfo.AddItem "Samplerate: " & .samplerate & " Hz"
        lstInfo.AddItem "Channels: " & .channels
        lstInfo.AddItem "Bits/Sample: " & .bitspersample
        lstInfo.AddItem "Bitrate: " & (.Bitrate / 1000) & " KBit/s"
        lstInfo.AddItem "Duration: " & (.Duration / 1000) & " s"
        lstInfo.AddItem ""
        lstInfo.AddItem "Tags:"

        For Each clsTag In .Tags
            lstInfo.AddItem clsTag.TagName & ": " & clsTag.TagValue

            Select Case LCase$(clsTag.TagName)
                Case strArtist
                    txtArtist.Text = clsTag.TagValue
                Case strAlbum
                    txtAlbum.Text = clsTag.TagValue
                Case strTitle
                    txtTitle.Text = clsTag.TagValue
                Case strTitleAlt
                    txtTitle.Text = clsTag.TagValue
            End Select
        Next
    End With

    OpenFile = True
End Function

Private Sub IDSPCallback_Samples( _
    intSamples() As Integer, _
    ByVal datalength As Long, _
    ByVal channels As Integer _
)

    '
End Sub

Private Function Min(val1 As Long, val2 As Long) As Long
    Min = IIf(val1 < val2, val1, val2)
End Function
