VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "VB Memcap"
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   474
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5130
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAllocate 
         Caption         =   "&Allocate"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuControl 
      Caption         =   "&Control"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Display"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "&Format"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSource 
         Caption         =   "S&ource"
      End
      Begin VB.Menu mnuCompression 
         Caption         =   "Co&mpression"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "&Select"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuScale 
         Caption         =   "Sc&ale"
         Checked         =   -1  'True
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlwaysVisible 
         Caption         =   "Al&ways Visible"
         Shortcut        =   ^W
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*
'* Author: E. J. Bantz Jr.
'* Copyright: None, use and distribute freely ...
'* E-Mail: ej@bantz.com
'* Web: http://ej.bantz.com
'*
Option Explicit

Private Sub Form_Load()
    
    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim Caps As CAPDRIVERCAPS
        
    '//Create Capture Window
    capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
    lwndC = capCreateCaptureWindowA(lpszName, WS_CAPTION Or WS_THICKFRAME Or WS_VISIBLE Or WS_CHILD, 0, 0, 160, 120, Me.hWnd, 0)

    '// Set title of window to name of driver
    SetWindowText lwndC, lpszName
    
    '// Set the video stream callback function
    capSetCallbackOnStatus lwndC, AddressOf MyStatusCallback
    capSetCallbackOnError lwndC, AddressOf MyErrorCallback
    
    '// Connect the capture window to the driver
    If capDriverConnect(lwndC, 0) Then
        '/////
        '// Only do the following if the connect was successful.
        '// if it fails, the error will be reported in the call
        '// back function.
        '/////
        '// Get the capabilities of the capture driver
        capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
        
        '// If the capture driver does not support a dialog, grey it out
        '// in the menu bar.
        If Caps.fHasDlgVideoSource = 0 Then mnuSource.Enabled = False
        If Caps.fHasDlgVideoFormat = 0 Then mnuFormat.Enabled = False
        If Caps.fHasDlgVideoDisplay = 0 Then mnuDisplay.Enabled = False
        
        '// Turn Scale on
        capPreviewScale lwndC, True
            
        '// Set the preview rate in milliseconds
        capPreviewRate lwndC, 66
        
        '// Start previewing the image from the camera
        capPreview lwndC, True
            
        '// Resize the capture window to show the whole image
        ResizeCaptureWindow lwndC

    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

    '// Disable all callbacks
    capSetCallbackOnError lwndC, vbNull
    capSetCallbackOnStatus lwndC, vbNull
    capSetCallbackOnYield lwndC, vbNull
    capSetCallbackOnFrame lwndC, vbNull
    capSetCallbackOnVideoStream lwndC, vbNull
    capSetCallbackOnWaveStream lwndC, vbNull
    capSetCallbackOnCapControl lwndC, vbNull
    

End Sub

Private Sub mnuAllocate_Click()

 Dim sFile As String * 250
 Dim lSize As Long
 
 '// Setup swap file for capture
 lSize = 1000000
 sFile = "C:\TEMP.AVI"
 capFileSetCaptureFile lwndC, sFile
 capFileAlloc lwndC, lSize
 
End Sub

Private Sub mnuAlwaysVisible_Click()
    
    mnuAlwaysVisible.Checked = Not (mnuAlwaysVisible.Checked)
    
    If mnuAlwaysVisible.Checked Then
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    Else
        SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If


End Sub

Private Sub mnuCompression_Click()
'   /*
'   * Display the Compression dialog when "Compression" is selected from
'   * the menu bar.
'   */
    
    capDlgVideoCompression lwndC

End Sub

Private Sub mnuCopy_Click()

    capEditCopy lwndC
        
End Sub

Private Sub mnuDisplay_Click()
'   /*
'   * Display the Video Display dialog when "Display" is selected from
'   * the menu bar.
'   */

    capDlgVideoDisplay lwndC
    
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Sub mnuFormat_Click()
'  /*
'   * Display the Video Format dialog when "Format" is selected from the
'   * menu bar.
'   */

    capDlgVideoFormat lwndC
    ResizeCaptureWindow lwndC

End Sub

Private Sub mnuPreview_Click()

    frmMain.StatusBar.SimpleText = vbNullString
    mnuPreview.Checked = Not (mnuPreview.Checked)
    capPreview lwndC, mnuPreview.Checked
    
End Sub

Private Sub mnuScale_Click()
    
    mnuScale.Checked = Not (mnuScale.Checked)
    capPreviewScale lwndC, mnuScale.Checked
    
    If mnuScale.Checked Then
       SetWindowLong lwndC, GWL_STYLE, WS_THICKFRAME Or WS_CAPTION Or WS_VISIBLE Or WS_CHILD
    Else
       SetWindowLong lwndC, GWL_STYLE, WS_BORDER Or WS_CAPTION Or WS_VISIBLE Or WS_CHILD
    End If

    ResizeCaptureWindow lwndC
    
End Sub

Private Sub mnuSelect_Click()
    
    frmSelect.Show vbModal, Me

End Sub

Private Sub mnuSource_Click()
'   /*
'    * Display the Video Source dialog when "Source" is selected from the
'    * menu bar.
'    */
    
    capDlgVideoSource lwndC

End Sub

Private Sub mnuStart_Click()
' /*
'  * If Start is selected from the menu, start Streaming capture.
'  * The streaming capture is terminated when the Escape key is pressed
'  */
    
    Dim sFileName As String
    Dim CAP_PARAMS As CAPTUREPARMS
    
    capCaptureGetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
    
    CAP_PARAMS.dwRequestMicroSecPerFrame = (1 * (10 ^ 6)) / 30  ' 30 Frames per second
    CAP_PARAMS.fMakeUserHitOKToCapture = True
    CAP_PARAMS.fCaptureAudio = False
    
    capCaptureSetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
    
    sFileName = "C:\myvideo.avi"
    
    capCaptureSequence lwndC  ' Start Capturing!
    capFileSaveAs lwndC, sFileName  ' Copy video from swap file into a real file.

End Sub
