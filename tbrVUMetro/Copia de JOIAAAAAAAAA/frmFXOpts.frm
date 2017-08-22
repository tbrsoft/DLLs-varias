VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFXOpts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "effects settings"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picAmpl 
      Height          =   3090
      Left            =   270
      ScaleHeight     =   3030
      ScaleWidth      =   5985
      TabIndex        =   28
      Top             =   480
      Width           =   6045
      Begin ComctlLib.Slider sldAmplDB 
         Height          =   390
         Left            =   930
         TabIndex        =   30
         Top             =   30
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   688
         _Version        =   327682
         Min             =   -12
         Max             =   12
      End
      Begin VB.Label lblAmplDB 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4620
         TabIndex        =   31
         Top             =   90
         Width           =   180
      End
      Begin VB.Label lblAmplBy 
         Caption         =   "Amplify by:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   29
         Top             =   60
         Width           =   840
      End
   End
   Begin VB.PictureBox picEQ 
      Height          =   3090
      Left            =   240
      ScaleHeight     =   3030
      ScaleWidth      =   6165
      TabIndex        =   23
      Top             =   480
      Width           =   6225
      Begin VB.PictureBox picEQAmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1935
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   179
         TabIndex        =   25
         Top             =   225
         Width           =   2715
      End
      Begin ComctlLib.Slider sldEQ 
         Height          =   1515
         Index           =   0
         Left            =   150
         TabIndex        =   24
         Top             =   825
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   2672
         _Version        =   327682
         Orientation     =   1
         Min             =   -12
         Max             =   12
         SelStart        =   -12
         TickStyle       =   2
         Value           =   -12
      End
      Begin VB.Label lbl12DB 
         AutoSize        =   -1  'True
         Caption         =   "[-12;12] dB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   450
         TabIndex        =   27
         Top             =   300
         Width           =   690
      End
      Begin VB.Label lblEQID 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   26
         Top             =   2325
         Width           =   105
      End
   End
   Begin VB.PictureBox picReverb 
      Height          =   3090
      Left            =   210
      ScaleHeight     =   3030
      ScaleWidth      =   6165
      TabIndex        =   0
      Top             =   450
      Width           =   6225
      Begin ComctlLib.Slider sldReverbLen 
         Height          =   405
         Left            =   1770
         TabIndex        =   1
         Top             =   390
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   714
         _Version        =   327682
         Min             =   1
         Max             =   1000
         SelStart        =   1
         TickStyle       =   3
         TickFrequency   =   100
         Value           =   1
      End
      Begin ComctlLib.Slider sldReverbAmp 
         Height          =   405
         Left            =   1770
         TabIndex        =   2
         Top             =   1170
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   714
         _Version        =   327682
         Min             =   1
         Max             =   9
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label lblReverbLen 
         AutoSize        =   -1  'True
         Caption         =   "Length (ms):"
         Height          =   195
         Left            =   780
         TabIndex        =   6
         Top             =   420
         Width           =   870
      End
      Begin VB.Label lblReverbFactor 
         Caption         =   "Amplifier:"
         Height          =   195
         Left            =   780
         TabIndex        =   5
         Top             =   1185
         Width           =   915
      End
      Begin VB.Label lbl1000ms 
         AutoSize        =   -1  'True
         Caption         =   "1000 ms"
         Height          =   195
         Left            =   4650
         TabIndex        =   4
         Top             =   765
         Width           =   600
      End
      Begin VB.Label lbl1ms 
         AutoSize        =   -1  'True
         Caption         =   "1 ms"
         Height          =   195
         Left            =   1875
         TabIndex        =   3
         Top             =   765
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   390
      Left            =   240
      TabIndex        =   21
      Top             =   4050
      Width           =   1515
   End
   Begin VB.PictureBox picShift 
      Height          =   3090
      Left            =   180
      ScaleHeight     =   3030
      ScaleWidth      =   6165
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   6225
      Begin MSComctlLib.Slider sldShDry 
         Height          =   270
         Left            =   1530
         TabIndex        =   8
         Top             =   255
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   476
         _Version        =   393216
         Min             =   -10
         SelStart        =   10
         TickStyle       =   3
         Value           =   10
      End
      Begin MSComctlLib.Slider sldShWet 
         Height          =   270
         Left            =   1530
         TabIndex        =   9
         Top             =   570
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   476
         _Version        =   393216
         Min             =   -10
         SelStart        =   10
         TickStyle       =   3
         Value           =   10
      End
      Begin MSComctlLib.Slider sldShFb 
         Height          =   270
         Left            =   1530
         TabIndex        =   10
         Top             =   885
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   476
         _Version        =   393216
         Min             =   -9
         Max             =   9
         SelStart        =   5
         TickStyle       =   3
         Value           =   5
      End
      Begin MSComctlLib.Slider sldShSwpRate 
         Height          =   270
         Left            =   1530
         TabIndex        =   11
         Top             =   1200
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   476
         _Version        =   393216
         Min             =   1
         Max             =   15
         SelStart        =   10
         TickStyle       =   3
         Value           =   10
      End
      Begin MSComctlLib.Slider sldShSwpRange 
         Height          =   270
         Left            =   1530
         TabIndex        =   12
         Top             =   1515
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   476
         _Version        =   393216
         Min             =   30
         Max             =   60
         SelStart        =   40
         TickStyle       =   3
         Value           =   40
      End
      Begin MSComctlLib.Slider sldShFreq 
         Height          =   270
         Left            =   1530
         TabIndex        =   13
         Top             =   1830
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   476
         _Version        =   393216
         Min             =   -10
         Max             =   150
         SelStart        =   100
         TickStyle       =   3
         Value           =   100
      End
      Begin VB.Label lblShDry 
         Caption         =   "Dry:"
         Height          =   195
         Left            =   270
         TabIndex        =   20
         Top             =   255
         Width           =   555
      End
      Begin VB.Label lblShWet 
         Caption         =   "Wet:"
         Height          =   195
         Left            =   270
         TabIndex        =   19
         Top             =   570
         Width           =   465
      End
      Begin VB.Label lblShFb 
         Caption         =   "Feedback:"
         Height          =   195
         Left            =   270
         TabIndex        =   18
         Top             =   885
         Width           =   825
      End
      Begin VB.Label lblShSwpRange 
         AutoSize        =   -1  'True
         Caption         =   "Sweep Range:"
         Height          =   195
         Left            =   270
         TabIndex        =   17
         Top             =   1515
         Width           =   1065
      End
      Begin VB.Label lblSwpRt 
         AutoSize        =   -1  'True
         Caption         =   "Sweep Rate:"
         Height          =   195
         Left            =   270
         TabIndex        =   16
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label lblShFreq 
         Caption         =   "Frequency:"
         Height          =   195
         Left            =   270
         TabIndex        =   15
         Top             =   1830
         Width           =   1185
      End
      Begin VB.Label lblCodeProj 
         AutoSize        =   -1  'True
         Caption         =   "http://www.codeproject.com/cs/media/cswavplayfx.asp"
         Height          =   195
         Left            =   450
         TabIndex        =   14
         Top             =   2400
         Width           =   4035
      End
   End
   Begin MSComctlLib.TabStrip strip 
      Height          =   3540
      Left            =   90
      TabIndex        =   22
      Top             =   60
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   6244
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Echo"
            Key             =   "ECHO"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Phase Shift"
            Key             =   "SHIFT"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Equalizer"
            Key             =   "EQ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Amplifier"
            Key             =   "AMPL"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFXOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINT
    x                       As Double
    y                       As Double
End Type

Private Const Pi            As Double = 3.14159265358979

Private Const F             As Long = 5
Private Const POINTS        As Long = 10

Private Const COLOR_1       As Long = &H122BC
Private Const COLOR_2       As Long = &H22A8E1
Private Const COLOR_3       As Long = &HCC00

Private p()                 As POINT

Private clsDSP              As clsDSP

Public Sub ShowEx( _
    dsp As clsDSP, _
    owner As Form _
)

    Dim i   As Long

    Set clsDSP = dsp

    ReDim p(1 To clsDSP.EqualizerBands) As POINT
    strip_Click

    With clsDSP
        sldReverbAmp.value = .EchoAmp * 10
        sldReverbLen.value = .EchoLength

        sldShDry.value = .PhaseShiftDry * 10
        sldShWet.value = .PhaseShiftWet * 10
        sldShFb.value = .PhaseShiftFeedback * 10
        sldShFreq.value = .PhaseShiftFrequency
        sldShSwpRange.value = .PhaseShiftSweepRange * 10
        sldShSwpRate.value = .PhaseShiftSweepRate * 10

        sldAmplDB.value = .AmplifyDB

        sldAmplDB_Scroll

        UpdateEQ
    End With

    Me.Show vbModal, owner
End Sub

Private Sub UpdateEQ()
    Dim i   As Long

    sldEQ(0).Left = picEQ.ScaleWidth / clsDSP.EqualizerBands - sldEQ(0).Width
    lblEQID(0).Left = sldEQ(0).Left + (15 * Screen.TwipsPerPixelX)

    ' create equalizer sliders
    If sldEQ.UBound = 0 Then
        For i = 1 To clsDSP.EqualizerBands - 1
            Load sldEQ(i)

            With sldEQ(i)
                .Left = sldEQ(i - 1).Left + .Width + (20 * Screen.TwipsPerPixelX)
                .Visible = True
            End With

            Load lblEQID(i)

            With lblEQID(i)
                .Left = sldEQ(i).Left + (15 * Screen.TwipsPerPixelX)
                .Visible = True
            End With
        Next
    End If

    ' get values (frequency, gain) for each band slider
    For i = 0 To clsDSP.EqualizerBands - 1
        sldEQ(i).value = -clsDSP.EqualizerBandGainDB(i)

        If clsDSP.EqualizerBandFrequency(i) > 1000 Then
            lblEQID(i).Caption = Fix(clsDSP.EqualizerBandFrequency(i) / 1000) & " kHz"
        Else
            lblEQID(i).Caption = Fix(clsDSP.EqualizerBandFrequency(i)) & " Hz"
        End If
    Next
    
    DrawEQ
End Sub

Private Sub sldAmplDB_Click()
    sldAmplDB_Scroll
End Sub

Private Sub sldAmplDB_Scroll()
    clsDSP.AmplifyDB = sldAmplDB.value
    lblAmplDB.Caption = sldAmplDB.value & " dB"
End Sub

Private Sub strip_Click()
    Select Case strip.SelectedItem.Key
        Case "ECHO"
            picEQ.Visible = False
            picReverb.Visible = True
            picShift.Visible = False
            picAmpl.Visible = False
        Case "SHIFT"
            picEQ.Visible = False
            picReverb.Visible = False
            picShift.Visible = True
            picAmpl.Visible = False
        Case "EQ"
            picEQ.Visible = True
            picReverb.Visible = False
            picShift.Visible = False
            picAmpl.Visible = False
        Case "AMPL"
            picEQ.Visible = False
            picReverb.Visible = False
            picShift.Visible = False
            picAmpl.Visible = True
    End Select
End Sub

Private Sub sldReverbAmp_Click()
    sldReverbAmp_Scroll
End Sub

Private Sub sldReverbAmp_Scroll()
    clsDSP.EchoAmp = sldReverbAmp.value / 10
End Sub

Private Sub sldReverbLen_Click()
    sldReverbLen_Scroll
End Sub

Private Sub sldReverbLen_Scroll()
    clsDSP.EchoLength = sldReverbLen.value
End Sub

Private Sub sldShDry_Click()
    sldShDry_Scroll
End Sub

Private Sub sldShDry_Scroll()
    clsDSP.PhaseShiftDry = sldShDry.value / 10
End Sub

Private Sub sldShFb_Click()
    sldShFb_Scroll
End Sub

Private Sub sldShFb_Scroll()
    clsDSP.PhaseShiftFeedback = sldShFb.value / 10
End Sub

Private Sub sldShFreq_Click()
    sldShFreq_Scroll
End Sub

Private Sub sldShFreq_Scroll()
    clsDSP.PhaseShiftFrequency = sldShFreq.value
End Sub

Private Sub sldShSwpRange_Click()
    sldShSwpRange_Scroll
End Sub

Private Sub sldShSwpRange_Scroll()
    clsDSP.PhaseShiftSweepRange = sldShSwpRange.value / 10
End Sub

Private Sub sldShSwpRate_Click()
    sldShFb_Scroll
End Sub

Private Sub sldShSwpRate_Scroll()
    clsDSP.PhaseShiftSweepRate = sldShSwpRate.value / 10
End Sub

Private Sub sldShWet_Click()
    sldShWet_Scroll
End Sub

Private Sub sldShWet_Scroll()
    clsDSP.PhaseShiftWet = sldShWet.value / 10
End Sub

Private Sub sldEQ_Click( _
    Index As Integer _
)

    sldEQ_Scroll Index
End Sub

Private Sub sldEQ_MouseUp( _
    Index As Integer, _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single _
)

    If Button = 2 Then
        sldEQ(Index).value = 0
        sldEQ_Scroll Index
    End If
End Sub

Private Sub sldEQ_Scroll( _
    Index As Integer _
)

    clsDSP.EqualizerBandGainDB(Index) = -sldEQ(Index).value
    DrawEQ
End Sub

Private Sub DrawEQ()
    Dim x       As Long
    Dim y       As Long
    Dim i       As Long
    Dim drwh    As Long
    Dim drww    As Long
    Dim drws    As Long
    Dim pah     As Long

    With picEQAmp
        .AutoRedraw = True
        .Cls
        .ScaleMode = 3

        drwh = .ScaleHeight
        drww = .ScaleWidth

        .BackColor = &H2C1F18
    End With

    ' space between 2 points
    drws = (drww - 10) / clsDSP.EqualizerBands

    ' eq band points
    For i = 1 To clsDSP.EqualizerBands
        p(i).x = drws * i - 4
        p(i).y = (sldEQ(i - 1).value / sldEQ(0).max) * (drwh / 2) + (drwh / 2)
    Next

    picEQAmp.Cls

    ' Slider
    picEQAmp.ForeColor = &H42585E

    For i = 1 To clsDSP.EqualizerBands
        picEQAmp.Line (p(i).x, 2)-(p(i).x, drwh - 2)
    Next

    ' draw preamp level
    picEQAmp.ForeColor = &H839EA7
'    pah = (-sldPreamp.value / sldPreamp.max) * (drwh / 2) + (drwh / 2) - 0.5
'    picEQAmp.Line (0, pah)-(drww, pah)

    ' draw a smoothed line with cosine interpolation
    i = 1
    x = 0

    y = CosineInterpolate(p(1).y, p(2).y, (x / F) / drws)
    picEQAmp.PSet (p(1).x, y)

    For i = 1 To clsDSP.EqualizerBands - 1
        For x = 0 To drws * F
            y = CosineInterpolate(p(i).y, p(i + 1).y, (x / F) / drws) - 0.5
            picEQAmp.ForeColor = GetGradColor(drwh, y, COLOR_1, COLOR_2, COLOR_3)
            picEQAmp.Line -(p(i).x + (x / F), y)
        Next
    Next
End Sub

' http://astronomy.swin.edu.au/~pbourke/other/interpolation/index.html
Private Function CosineInterpolate( _
    ByVal Y1 As Double, _
    ByVal Y2 As Double, _
    ByVal mu As Double _
) As Double

    Dim mu2 As Double

    mu2 = (1 - Cos(mu * Pi)) / 2
    CosineInterpolate = (Y1 * (1 - mu2) + Y2 * mu2)
End Function
