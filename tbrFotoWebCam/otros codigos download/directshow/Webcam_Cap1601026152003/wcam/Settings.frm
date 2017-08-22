VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   6165
   ClientLeft      =   6345
   ClientTop       =   3870
   ClientWidth     =   4065
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   4065
   Begin VB.ComboBox cbMethod 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1260
      TabIndex        =   24
      ToolTipText     =   "Contains the available methods for image scan"
      Top             =   4680
      Width           =   2670
   End
   Begin VB.TextBox tTolerance 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   21
      ToolTipText     =   "Quality of the image in %"
      Top             =   4275
      Width           =   690
   End
   Begin VB.CheckBox cbMvOnly 
      Caption         =   "Moves Only"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1980
      TabIndex        =   19
      ToolTipText     =   "Take the picture only when a move/change  is detected"
      Top             =   3915
      Width           =   1860
   End
   Begin VB.CheckBox cbMDetect 
      Caption         =   "Motion Detect"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   18
      ToolTipText     =   "Enable / disable the motion detection"
      Top             =   3915
      Width           =   1770
   End
   Begin MSComctlLib.Slider slSlider1 
      Height          =   285
      Left            =   3150
      TabIndex        =   17
      ToolTipText     =   "Adjust image quality"
      Top             =   3420
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   503
      _Version        =   393216
      Min             =   10
      Max             =   100
      SelStart        =   10
      TickFrequency   =   10
      Value           =   10
   End
   Begin VB.CheckBox cbTimeStamp 
      Caption         =   "Timestamp"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   16
      ToolTipText     =   "Add a timestamp to the image"
      Top             =   630
      Width           =   2895
   End
   Begin VB.TextBox tQuality 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3150
      TabIndex        =   15
      ToolTipText     =   "Quality of the image in %"
      Top             =   3150
      Width           =   825
   End
   Begin VB.ComboBox cbFileType 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1260
      TabIndex        =   13
      ToolTipText     =   "File type"
      Top             =   3150
      Width           =   960
   End
   Begin VB.TextBox tFilePrefix 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1260
      TabIndex        =   11
      ToolTipText     =   "File prefix"
      Top             =   2745
      Width           =   2715
   End
   Begin VB.TextBox tFilePath 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Directory path"
      Top             =   2340
      Width           =   3885
   End
   Begin VB.CommandButton btBrowse 
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1980
      TabIndex        =   8
      ToolTipText     =   "Browse for the image directory"
      Top             =   1935
      Width           =   1995
   End
   Begin VB.OptionButton obModeS 
      Caption         =   "Image sequences"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1035
      TabIndex        =   6
      ToolTipText     =   "Use image sequences, add a time stamp to the image name"
      Top             =   1400
      Width           =   1725
   End
   Begin VB.CommandButton btSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "Save the settings"
      Top             =   5580
      Width           =   1635
   End
   Begin VB.CommandButton btCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2385
      TabIndex        =   3
      ToolTipText     =   "Cancel the settings"
      Top             =   5580
      Width           =   1635
   End
   Begin VB.OptionButton obModeR 
      Caption         =   "Replace image"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1035
      TabIndex        =   2
      ToolTipText     =   "Use a unique image per camera"
      Top             =   990
      Width           =   1725
   End
   Begin VB.TextBox tCapInt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2385
      TabIndex        =   1
      ToolTipText     =   "Capture interval in seconds"
      Top             =   225
      Width           =   690
   End
   Begin MSComctlLib.Slider slSlider2 
      Height          =   285
      Left            =   1980
      TabIndex        =   22
      ToolTipText     =   "Adjust tolerance"
      Top             =   4275
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   503
      _Version        =   393216
      Max             =   100
      TickFrequency   =   10
   End
   Begin VB.Label Label8 
      Caption         =   "Method :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   23
      Top             =   4680
      Width           =   1050
   End
   Begin VB.Label Label7 
      Caption         =   "Tolerance :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   20
      Top             =   4275
      Width           =   1050
   End
   Begin VB.Shape shFrame3 
      Height          =   1365
      Left            =   90
      Top             =   3780
      Width           =   3885
   End
   Begin VB.Shape shFrame1 
      Height          =   915
      Left            =   90
      Top             =   45
      Width           =   3885
   End
   Begin VB.Shape shFrame2 
      Height          =   960
      Left            =   90
      Top             =   945
      Width           =   3885
   End
   Begin VB.Label Label6 
      Caption         =   "Quality :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2250
      TabIndex        =   14
      Top             =   3150
      Width           =   870
   End
   Begin VB.Label Label5 
      Caption         =   "File type :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   12
      Top             =   3150
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "File prefix :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   10
      Top             =   2745
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Images directory :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   7
      Top             =   1935
      Width           =   1770
   End
   Begin VB.Label Label2 
      Caption         =   "Mode :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   5
      Top             =   990
      Width           =   870
   End
   Begin VB.Label Label1 
      Caption         =   "Capture interval (sec) :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   2220
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btBrowse_Click()
'
frmSeekDir.Show
End Sub

Private Sub btCancel_Click()
Call Form_Terminate
End Sub

Private Sub btSave_Click()
'
Dim sSplitV() As String
'
iCfgInterval = Val(Trim(Me.tCapInt.Text))
If Me.obModeR.Value Then
 iCfgMode = 0
End If
If Me.obModeS.Value Then
 iCfgMode = 1
End If
sCfgPath = Me.tFilePath.Text
sCfgFile = Me.tFilePrefix.Text
sCfgFileExt = Me.cbFileType.Text
iCfgFmtQuality = Val(Trim(Me.tQuality.Text))
iCfgTimeStamp = Me.cbTimeStamp.Value
iCfgMDSwitch = Me.cbMDetect.Value
iCfgMDMvOnly = Me.cbMvOnly.Value
iCfgMDTolerance = Val(Trim(Me.tTolerance.Text))
sSplitV = Split(Me.cbMethod.Text, " ", -1, vbBinaryCompare)
If UBound(sSplitV) > 0 Then
 iCfgMDMethod = Val(Trim(sSplitV(0)))
Else
 iCfgMDMethod = 0
End If
Call WriteCfgFile
Call Form_Terminate
End Sub

Private Sub cbFileType_Change()
'
If Me.cbFileType.Text = "BMP" Then
 Me.tQuality.Text = "100"
 Me.slSlider1.Enabled = False
End If
If Me.cbFileType.Text = "JPG" Then
 Me.slSlider1.Enabled = True
End If
End Sub

Private Sub cbFileType_Click()
Call cbFileType_Change
End Sub

Private Sub cbMDetect_Click()
'
Dim iretval As Integer
Dim sMethods() As String
Dim lidx As Long
'
If Me.cbMDetect.Value = 0 Then
 Me.cbMvOnly.Value = 0
 Me.cbMvOnly.Enabled = False
 Me.tTolerance.Text = ""
 Me.slSlider2.Enabled = False
 Me.cbMethod.Clear
 Me.cbMethod.Enabled = False
 Me.cbMethod.Text = ""
Else
 Me.cbMvOnly.Value = 0
 Me.cbMvOnly.Enabled = True
 Me.tTolerance.Text = "0"
 Me.slSlider2.Value = 0
 Me.slSlider2.Enabled = True
 Me.cbMethod.Enabled = True
 Set clsImgDiff = New ImgDiff
 ReDim sMethods(0)
 Me.cbMethod.Clear
 iretval = clsImgDiff.GetMethodList(sMethods)
 Set clsImgDiff = Nothing
 If UBound(sMethods) > 0 Then
  For lidx = 1 To UBound(sMethods)
   Me.cbMethod.AddItem (Trim(CStr(lidx)) & " : " & sMethods(lidx))
  Next lidx
  Me.cbMethod.Text = Trim(CStr(iCfgMDMethod)) & " : " & sMethods(iCfgMDMethod)
 Else
  Me.cbMethod.Text = ""
 End If
End If
End Sub

Private Sub Form_Load()
'
' Load form
'
Dim sMethods() As String
Dim iretval As Integer
Dim lidx As Long
'
Me.tCapInt.Text = CStr(iCfgInterval)
If iCfgMode = 0 Then
 Me.obModeR.Value = True
End If
If iCfgMode = 1 Then
 Me.obModeS.Value = True
End If
Me.tFilePath.Text = sCfgPath
'
Me.cbFileType.Clear
Me.cbFileType.AddItem ("BMP")
Me.cbFileType.AddItem ("JPG")
Me.cbFileType.Text = sCfgFileExt
'
Me.tFilePrefix.Text = sCfgFile
Me.tQuality.Text = CStr(iCfgFmtQuality)
Me.tQuality.Enabled = False
Me.slSlider1.Value = iCfgFmtQuality
If sCfgFileExt = "BMP" Then
 Me.slSlider1.Enabled = False
End If
Me.cbTimeStamp.Value = iCfgTimeStamp
Me.cbMDetect.Value = iCfgMDSwitch
Me.tTolerance.Enabled = False
If iCfgMDSwitch = 0 Then
 Me.cbMvOnly.Value = 0
 Me.cbMvOnly.Enabled = False
 Me.tTolerance.Text = ""
 Me.slSlider2.Enabled = False
 Me.cbMethod.Clear
 Me.cbMethod.Text = ""
 Me.cbMethod.Enabled = False
Else
 Me.cbMvOnly.Value = iCfgMDMvOnly
 Me.tTolerance.Text = CStr(iCfgMDTolerance)
 Me.slSlider2.Value = iCfgMDTolerance
 Set clsImgDiff = New ImgDiff
 ReDim sMethods(0)
 Me.cbMethod.Clear
 iretval = clsImgDiff.GetMethodList(sMethods)
 Set clsImgDiff = Nothing
 If UBound(sMethods) > 0 Then
  For lidx = 1 To UBound(sMethods)
   Me.cbMethod.AddItem (Trim(CStr(lidx)) & " : " & sMethods(lidx))
  Next lidx
  Me.cbMethod.Text = Trim(CStr(iCfgMDMethod)) & " : " & sMethods(iCfgMDMethod)
 Else
  Me.cbMethod.Text = ""
 End If
  
End If
End Sub

Private Sub Form_Resize()
'
Dim lMinWidth As Long
Dim lMinHeight As Long
Dim lFWidth As Long
Dim lFHeight As Long
'
lMinWidth = 4185
lMinHeight = 6180
If Me.width < lMinWidth Then
 Me.width = lMinWidth
End If
If Me.height < lMinHeight Then
 Me.height = lMinHeight
End If
'
lFWidth = Me.width
lFHeight = Me.height
'
Me.shFrame1.width = lFWidth - 300
'
Me.shFrame2.width = lFWidth - 300
'
Me.shFrame3.width = lFWidth - 300
'
Me.btBrowse.width = lFWidth - 2190
'
Me.tFilePath.width = lFWidth - 300
'
Me.tFilePrefix.width = lFWidth - 1470
'
Me.Label6.Left = lFWidth - 1935
'
Me.tQuality.Left = lFWidth - 1035
'
Me.slSlider1.Left = lFWidth - 1035
'
Me.cbMvOnly.width = lFWidth - 2205
'
Me.slSlider2.width = lFWidth - 2205
'
Me.cbMethod.width = lFWidth - 1515
'
Me.btSave.Top = lFHeight - 990
'
Me.btCancel.Top = lFHeight - 990
Me.btCancel.Left = lFWidth - 1800
End Sub

Private Sub Form_Terminate()
'
' Terminate
'
Dim frmCurrent As Form
Dim idx As Integer
'
' Close forms
'
For Each frmCurrent In Forms
 If frmCurrent.Caption = "Images Directory" Then
    If frmCurrent.Visible Then
     frmCurrent.Hide
    End If
    Unload frmCurrent
 End If
Next frmCurrent
Me.Hide
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Form_Terminate
End Sub



Private Sub slSlider1_Scroll()
'
Me.tQuality.Text = CStr(Me.slSlider1.Value)
End Sub

Private Sub slSlider2_Scroll()
Me.tTolerance.Text = CStr(Me.slSlider2.Value)
End Sub

Private Sub tCapInt_Click()
Call tCapInt_LostFocus
End Sub

Private Sub tCapInt_LostFocus()
'
Dim sCapInt As String
Dim iCapInt As Integer
'
sCapInt = Me.tCapInt.Text
If IsNumeric(Trim(sCapInt)) Then
 iCapInt = Val(Trim(sCapInt))
 If iCapInt < 20 Then
  Me.tCapInt.Text = "20"
 End If
 If iCapInt > 3600 Then
  Me.tCapInt.Text = "3600"
 End If
Else
 Me.tCapInt.Text = "60"
End If
End Sub

Private Sub tFilePrefix_Click()
Call tFilePrefix_LostFocus
End Sub

Private Sub tFilePrefix_LostFocus()
'
If Me.tFilePrefix.Text = "" Then
 Me.tFilePrefix.Text = "webcam"
End If
End Sub



