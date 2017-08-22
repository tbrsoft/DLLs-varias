Attribute VB_Name = "MemCap"
'Aqui estan los procedimientos que son llamados ante algun evento del driver
'estos deben estar obligatoriamente en modulos BAS
'defino aqui una instancia del tbrWebCam para avisarle a el

'*
'* Author: E. J. Bantz Jr.
'* Copyright: None, use and distribute freely ...
'* E-Mail: ejbantz@usa.net
'* Web: http://www.inlink.com/~ejbantz

'// ------------------------------------------------------------------
'//  Windows API Constants / Types / Declarations
'// ------------------------------------------------------------------
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = 1
Public Const SWP_NOZORDER = &H4
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SM_CYCAPTION = 4
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const WS_EX_TRANSPARENT = &H20&
Public Const GWL_STYLE = (-16)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


'// Memory manipulation
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Declare Function lStrCpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As Any, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (ByVal hpvDest As Long, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub hmemcpy Lib "kernel32" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
    
'// Window manipulation
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Const WM_SETTEXT = &HC 'establece texto en un textbox

Public HwndMSGS As Long

Function MyFrameCallback(ByVal lwnd As Long, ByVal lpVHdr As Long) As Long

    'Debug.Print "FrameCallBack"
    
    Dim VideoHeader As VIDEOHDR
    Dim VideoData() As Byte
    
    '//Fill VideoHeader with data at lpVHdr
    RtlMoveMemory VarPtr(VideoHeader), lpVHdr, Len(VideoHeader)
    
    '// Make room for data
    ReDim VideoData(VideoHeader.dwBytesUsed)
    
    '//Copy data into the array
    RtlMoveMemory VarPtr(VideoData(0)), VideoHeader.lpData, VideoHeader.dwBytesUsed

    'lo saque yo por que me parece que jodia el debug abajo
    'Debug.Print VideoHeader.dwBytesUsed
    'Debug.Print VideoData
    
End Function

'Function MyYieldCallback(lwnd As Long) As Long
'
'    Debug.Print "Yield"
'
'End Function

Function MyErrorCallback(ByVal lwnd As Long, ByVal iID As Long, ByVal ipstrStatusText As Long) As Long
    
    If iID = 0 Then Exit Function
    
    Dim sStatusText As String
    Dim usStatusText As String
    
    'Convert the Pointer to a real VB String
    sStatusText = String$(255, 0)                                      '// Make room for message
    lstrcpy StrPtr(sStatusText), ipstrStatusText                       '// Copy message into String
    sStatusText = Left$(sStatusText, InStr(sStatusText, Chr$(0)) - 1)  '// Only look at left of null
    usStatusText = StrConv(sStatusText, vbUnicode)                     '// Convert Unicode
            
    SendMessageS HwndMSGS, WM_SETTEXT, 0, ByVal "Error: " + sStatusText

End Function

Function MyStatusCallback(ByVal lwnd As Long, ByVal iID As Long, ByVal ipstrStatusText As Long) As Long

    If iID = 0 Then Exit Function
   
    Dim sStatusText As String
    Dim usStatusText As String
    
    '// Convert the Pointer to a real VB String
    sStatusText = String$(255, 0)                                      '// Make room for message
    lstrcpy StrPtr(sStatusText), ipstrStatusText                       '// Copy message into String
    sStatusText = Left$(sStatusText, InStr(sStatusText, Chr$(0)) - 1)  '// Only look at left of null
    usStatusText = StrConv(sStatusText, vbUnicode)                     '// Convert Unicode

    SendMessageS HwndMSGS, WM_SETTEXT, 0, ByVal "STAT: " + usStatusText
    
End Function

Sub ResizeCaptureWindow(ByVal lwnd As Long)

    Dim CAPSTATUS As CAPSTATUS
    Dim lCaptionHeight As Long
    Dim lX_Border As Long
    Dim lY_Border As Long
    
    
    lCaptionHeight = GetSystemMetrics(SM_CYCAPTION)
    lX_Border = GetSystemMetrics(SM_CXFRAME)
    lY_Border = GetSystemMetrics(SM_CYFRAME)
    
    '// Get the capture window attributes .. width and height
    If capGetStatus(lwnd, VarPtr(CAPSTATUS), Len(CAPSTATUS)) Then
        
        Dim myCX As Long
        myCX = CAPSTATUS.uiImageWidth + (lX_Border * 2)
        Dim myCY As Long
        myCY = CAPSTATUS.uiImageHeight + lCaptionHeight + (lY_Border * 2)
        '// Resize the capture window to the capture sizes
        'si devuelve cero es error
        If SetWindowPos(lwnd, HWND_BOTTOM, 0, 0, myCX, myCY, SWP_NOMOVE Or SWP_NOZORDER) = 0 Then
            SendMessageS HwndMSGS, WM_SETTEXT, 0, "Resize no puede"
        Else
            SendMessageS HwndMSGS, WM_SETTEXT, 0, "Resize: " + CStr(myCX) + "," + CStr(myCY)
        End If
        
    Else
        SendMessageS HwndMSGS, WM_SETTEXT, 0, "Resize mal"
    End If

End Sub

Public Sub UpdateStatusCaps()
    
    
    Dim CAPSTATUS As CAPSTATUS, ST As String
    If capGetStatus(lwnd, VarPtr(CAPSTATUS), Len(CAPSTATUS)) Then
    
        ST = CStr(CAPSTATUS.uiImageWidth)
    
    '    uiImageWidth As Long                    '// Width of the image
    '    uiImageHeight As Long                   '// Height of the image
    '    fLiveWindow As Long                     '// Now Previewing video?
    '    fOverlayWindow As Long                  '// Now Overlaying video?
    '    fScale As Long                          '// Scale image to client?
    '    ptScroll As POINTAPI                    '// Scroll position
    '    fUsingDefaultPalette As Long            '// Using default driver palette?
    '    fAudioHardware As Long                  '// Audio hardware present?
    '    fCapFileExists As Long                  '// Does capture file exist?
    '    dwCurrentVideoFrame As Long             '// # of video frames cap'td
    '    dwCurrentVideoFramesDropped As Long     '// # of video frames dropped
    '    dwCurrentWaveSamples As Long            '// # of wave samples cap'td
    '    dwCurrentTimeElapsedMS As Long          '// Elapsed capture duration
    '    hPalCurrent As Long                     '// Current palette in use
    '    fCapturingNow As Long                   '// Capture in progress?
    '    dwReturn As Long                        '// Error value after any operation
    '    wNumVideoAllocated As Long              '// Actual number of video buffers
    '    wNumAudioAllocated As Long              '// Actual number of audio buffers
    
    
    End If
End Sub

Function MyVideoStreamCallback(lwnd As Long, lpVHdr As Long) As Long
   'como parece no necesario lo saque
    Beep  '// Replace this with your code!

End Function

'Function MyWaveStreamCallback(lwnd As Long, lpVHdr As Long) As Long
'
'    Debug.Print "WaveStream"
'
'End Function

