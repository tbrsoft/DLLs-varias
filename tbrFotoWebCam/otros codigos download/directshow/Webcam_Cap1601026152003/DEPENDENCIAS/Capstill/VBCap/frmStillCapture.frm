VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStillCapture 
   Caption         =   "Still Image Capture"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   5520
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "Render File"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Effect 
      Caption         =   "Snap + Effect"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3015
      Left            =   840
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   7
      Top             =   3480
      Width           =   4215
   End
   Begin VB.CommandButton Grab 
      Caption         =   "Snap to Mem"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtImageFile 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Text            =   "StillImage.bmp"
      Top             =   2640
      Width           =   3975
   End
   Begin VB.ListBox lstFilters 
      Height          =   1815
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton Snap 
      Caption         =   "&Snap File"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Preview 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Still Filename"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select Source Filter then press Preview"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmStillCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gGraph As IMediaControl
Dim gRegFilters As Object
Dim gCapStill As VBGrabber

' GDI functions to draw a DIBSection into a DC
Private Declare Function CreateCompatibleDC Lib "GDI32" _
    (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "GDI32" _
    (ByVal hdc As Long, ByVal hbitmap As Long) As Long
Private Declare Function BitBlt Lib "GDI32" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal width As Long, ByVal height As Long, _
    ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal mode As Long) _
    As Long
Private Declare Sub DeleteDC Lib "GDI32" _
    (ByVal hdc As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (dest As Any, src As Any, ByVal count As Long)
    
' non-portable (win32 only) types and functions to
' convert a bitmap into a safe array of bytes
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgsabound(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" _
    (Ptr() As Any) As Long
    
    
Private Sub btnOpen_Click()
    On Error GoTo cancelopen
    dlgFile.ShowOpen
    On Error GoTo 0
    
    ' make a new graph
    Set gGraph = Nothing
    Set gCapStill = Nothing
    Set gGraph = New FilgraphManager
    Set gRegFilters = gGraph.RegFilterCollection
    
    
    ' add the grabber including vb wrapper and default props
    Dim filter As IRegFilterInfo
    Dim fGrab As IFilterInfo
    For Each filter In gRegFilters
        If filter.Name = "SampleGrabber" Then
            filter.filter fGrab
            
            ' wrap this filter in the capstill vb wrapper
            ' also sets rgb-24 media type and other properties
            Set gCapStill = New VBGrabber
            gCapStill.FilterInfo = fGrab
            Exit For
        End If
    Next filter
    
    Dim fSrc As IFilterInfo
    gGraph.AddSourceFilter dlgFile.FileName, fSrc
    
    ' find first output on src
    Dim pinOut As IPinInfo
    For Each pinOut In fSrc.Pins
        If pinOut.Direction = 1 Then
            Exit For
        End If
    Next pinOut
    
    ' find first input on grabber and connect
    Dim pinIn As IPinInfo
    For Each pinIn In fGrab.Pins
        If pinIn.Direction = 0 Then
            pinOut.Connect pinIn
            Exit For
        End If
    Next pinIn
    
    ' find grabber output pin and render
    For Each pinOut In fGrab.Pins
        If pinOut.Direction = 1 Then
            pinOut.Render
            Exit For
        End If
    Next pinOut
    
    
    ' run graph and we are successfully in preview mode
    gGraph.Run
    
    
cancelopen:

End Sub

' Demonstrates capturing a still image to memory and then
' accessing the bits directly using a highly non-portable technique
Private Sub Effect_Click()
    Dim bma As IBitmapAccess
    Set bma = gCapStill.CapToMem
        
    ' this highly-non portable hack gets you access to the bits as a
    ' two dimensional array by creating a SAFEARRAY structure and
    ' copying it over a properly declared array
    Dim sa As SAFEARRAY
    sa.cbElements = 1
    sa.cDims = 2
    sa.rgsabound(0).lLbound = 0
    sa.rgsabound(0).cElements = bma.height
    sa.rgsabound(1).lLbound = 0
    sa.rgsabound(1).cElements = bma.Stride
    sa.pvData = bma.bits
    
    ' this array points to nothing
    Dim bits() As Byte
    
    'set array to point to the safearray info
    CopyMemory ByVal VarPtrArray(bits()), VarPtr(sa), 4
    
    ' now we can access the bitmap as bits(x, y)
    For x = 0 To bma.Stride - 3 Step 3
        For y = 0 To bma.height - 1
            bits(x, y) = (bits(x, y) + 60) Mod 255
        Next y
    Next x
    
    'clean up array hack :set array back to point to nothing
    CopyMemory ByVal VarPtrArray(bits()), 0&, 4
    
    ShowBitmap bma
End Sub

Private Sub Form_Load()
    Set gGraph = New FilgraphManager
    Set gRegFilters = gGraph.RegFilterCollection
    RefreshRegFilters
End Sub


Private Sub RefreshRegFilters()
' update the listbox of registered filters
' using the global variable gRegFilters
    Dim filter As IRegFilterInfo
    lstFilters.Clear
    If Not gRegFilters Is Nothing Then
        For Each filter In gRegFilters
            lstFilters.AddItem filter.Name
        Next filter
    End If
    If lstFilters.ListCount > 0 Then
        lstFilters.ListIndex = 0  ' select first in list
    End If
End Sub

Private Sub Grab_Click()
    Dim bma As IBitmapAccess
    Set bma = gCapStill.CapToMem
    
    ShowBitmap bma
End Sub

Private Sub lstFilters_DblClick()
    Preview_Click
End Sub

Private Sub Preview_Click()
    ' make a new graph
    Set gGraph = Nothing
    Set gCapStill = Nothing
    Set gGraph = New FilgraphManager
    Set gRegFilters = gGraph.RegFilterCollection
    
    
    ' add the grabber including vb wrapper and default props
    Dim filter As IRegFilterInfo
    Dim fGrab As IFilterInfo
    For Each filter In gRegFilters
        If filter.Name = "SampleGrabber" Then
            filter.filter fGrab
            
            ' wrap this filter in the capstill vb wrapper
            ' also sets rgb-24 media type and other properties
            Set gCapStill = New VBGrabber
            gCapStill.FilterInfo = fGrab
            Exit For
        End If
    Next filter
    
    ' add the selected source filter
    Dim fSrc As IFilterInfo
    For Each filter In gRegFilters
        If filter.Name = lstFilters.Text Then
            filter.filter fSrc
            Exit For
        End If
    Next filter
    
    ' check for crossbar and select decoder
    Dim xbar As CrossbarInfo
    Set xbar = New CrossbarInfo
    On Error GoTo NoXBar
    xbar.SetFilter fSrc
    
    Dim idx As Long
    For idx = 0 To xbar.Inputs - 1
        Dim pin As String
        pin = xbar.Name(True, idx)
        ' probably you want a dialog listing all input pins
        ' that xbar.CanRoute to the out
        ' or something hardwired like:
        'If pin = "1: Video Composite In" Then
            'xbar.Route idx, 0
            'Exit For
        'End If
    Next idx
    
    If xbar.Standard <> AnalogVideo_PAL_B Then
        xbar.Standard = AnalogVideo_PAL_B
    End If
    
NoXBar:
    On Error Resume Next
        
    ' find first output on src
    Dim pinOut As IPinInfo
    For Each pinOut In fSrc.Pins
        If pinOut.Direction = 1 Then
            Exit For
        End If
    Next pinOut
    
    'restore specified file before dlg
    Dim pSC As StreamConfig
    Set pSC = New StreamConfig
    pSC.pin = pinOut
    If pSC.SupportsConfig Then
        If Dir$("mtsave.mt") <> "" Then
            pSC.Restore ("mtsave.mt")
        End If
    End If
    
    ' show format of output pin before rendering
    Dim ppropOut As PinPropInfo
    Set ppropOut = New PinPropInfo
    ppropOut.pin = pinOut
    ppropOut.ShowPropPage 0
            
    ' save selected format to file
    If pSC.SupportsConfig Then
        pSC.SaveCurrentFormat ("mtsave.mt")
    End If
        
                
    ' find first input on grabber and connect
    Dim pinIn As IPinInfo
    For Each pinIn In fGrab.Pins
        If pinIn.Direction = 0 Then
            pinOut.Connect pinIn
            Exit For
        End If
    Next pinIn
    
    ' find grabber output pin and render
    For Each pinOut In fGrab.Pins
        If pinOut.Direction = 1 Then
            pinOut.Render
            Exit For
        End If
    Next pinOut
    
    
    ' run graph and we are successfully in preview mode
    gGraph.Run
End Sub

Private Sub Snap_Click()
    gCapStill.FileName = txtImageFile.Text
    gCapStill.CaptureStill
End Sub


Public Sub ShowBitmap(bma As IBitmapAccess)
    ' set correct size of image and then
    ' BitBlt to the picture control's HDC
    
    Picture1.width = bma.width * Screen.TwipsPerPixelX
    Picture1.height = bma.height * Screen.TwipsPerPixelY
    
    Dim hbm As Long
    hbm = bma.DIBSection
    
    Dim hMemDc As Long
    hMemDc = CreateCompatibleDC(Picture1.hdc)
    Dim hOldBM As Long
    hOldBM = SelectObject(hMemDc, hbm)
    BitBlt Picture1.hdc, _
                0, 0, bma.width, bma.height, _
                hMemDc, 0, 0, &HCC0020
    SelectObject hMemDc, hOldBM
    DeleteDC hMemDc
    ' hbm is owned by the BMA object
    
    Picture1.Refresh
End Sub
