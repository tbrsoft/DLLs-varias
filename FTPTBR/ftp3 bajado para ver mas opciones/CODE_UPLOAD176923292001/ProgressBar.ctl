VERSION 5.00
Begin VB.UserControl ProgressBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ======================================================================================
' cProgBar control
' Steve McMahon
' 02 June 1998
'
' A simple implementation of the Common Control Progress Bar
' ======================================================================================

' ======================================================================================
' API declares:
' ======================================================================================

' Memory functions:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
' Window functions
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
' Window style bit functions:
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
    ) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long _
    ) As Long
' Window Long indexes:
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_HINSTANCE = (-6)
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_ID = (-12)
Private Const GWL_STYLE = (-16)
Private Const GWL_USERDATA = (-21)
Private Const GWL_WNDPROC = (-4)
' Style:
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000

 ' Window relationship functions:
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
' WIndow position:
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const HWND_NOTOPMOST = -2
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
' Messages
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_USER = &H400

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

' common controls:
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

' progress bar:
Private Const PROGRESS_CLASSA = "msctls_progress32"

'Style
Private Const PBS_SMOOTH = &H1
Private Const PBS_VERTICAL = &H4
Private Const PBM_SETRANGE = (WM_USER + 1)
Private Const PBM_SETPOS = (WM_USER + 2)
Private Const PBM_DELTAPOS = (WM_USER + 3)
Private Const PBM_SETSTEP = (WM_USER + 4)
Private Const PBM_STEPIT = (WM_USER + 5)
Private Const PBM_SETRANGE32 = (WM_USER + 6)
Private Const PBM_GETRANGE = (WM_USER + 7)
Private Const PBM_GETPOS = (WM_USER + 8)
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR

Private Type PPBRange
   iLow As Long
   iHigh As Long
End Type


' ======================================================================================
' Implementation:
' ======================================================================================
Public Enum EPBBorderStyle
    epbBorderStyleNone
    epbBorderStyleSingle
    epdBorderStyle3d
End Enum
Public Enum EPBOrientation
    epbHorizontal
    epbVertical
End Enum

' ======================================================================================
' Private variables:
' ======================================================================================
Private m_hWnd As Long
'Private m_oBackColor As OLE_COLOR
'Private m_oForeColor As OLE_COLOR
Private m_bSmooth As Boolean
Private m_eOrientation As EPBOrientation
Private m_eBorderStyle As EPBBorderStyle
Private m_lPosition As Long
Private m_lMin As Long
Private m_lMax As Long
Private m_lStep As Long

Public Property Get Orientation() As EPBOrientation
   Orientation = m_eOrientation
End Property
Public Property Let Orientation(ByVal eOrientation As EPBOrientation)
   If (m_eOrientation <> eOrientation) Then
      m_eOrientation = eOrientation
      If (m_hWnd <> 0) Then
         ' set style...
         pRecreate
      End If
      PropertyChanged "Orientation"
   End If
End Property
Public Property Get Min() As Long
   Min = m_lMin
End Property
Public Property Let Min(ByVal lMin As Long)
   If (m_lMin <> lMin) Then
      m_lMin = lMin
      If (m_hWnd <> 0) Then
         pSetRange
      End If
      PropertyChanged "Min"
   End If
End Property
Public Property Get Max() As Long
   Max = m_lMax
End Property
Public Property Let Max(ByVal lMax As Long)
   If (m_lMax <> lMax) Then
      m_lMax = lMax
      If (m_hWnd <> 0) Then
         pSetRange
      End If
      PropertyChanged "Max"
   End If
End Property
Public Property Let Smooth(ByVal bSmooth As Boolean)
Dim lStyle As Long
Dim hP As Long
   If (m_bSmooth <> bSmooth) Then
      m_bSmooth = bSmooth
      If (m_hWnd <> 0) Then
         ' set style..
         pRecreate
      End If
      PropertyChanged "Smooth"
   End If
End Property
Public Property Get Smooth() As Boolean
   Smooth = m_bSmooth
End Property
Public Property Get hWnd() As Long
   hWnd = m_hWnd
End Property
Public Property Get BorderStyle() As EPBBorderStyle
   BorderStyle = m_eBorderStyle
End Property
Property Let BorderStyle(ByVal eBorderStyle As EPBBorderStyle)
Dim lStyle As Long
Dim lCStyle As Long
   If (m_eBorderStyle <> eBorderStyle) Then
      m_eBorderStyle = eBorderStyle
      lStyle = GetWindowLong(UserControl.hWnd, GWL_EXSTYLE)
      If (m_hWnd <> 0) Then
         lCStyle = GetWindowLong(m_hWnd, GWL_EXSTYLE)
      End If
      If (eBorderStyle <> epdBorderStyle3d) Then
         lStyle = lStyle And Not WS_EX_CLIENTEDGE
         If (eBorderStyle = epbBorderStyleSingle) Then
            lCStyle = lCStyle Or WS_EX_STATICEDGE
         Else
            lCStyle = lCStyle And Not WS_EX_STATICEDGE
         End If
      Else
         lStyle = lStyle Or WS_EX_CLIENTEDGE
         lCStyle = lCStyle And Not WS_EX_STATICEDGE
      End If
      If (m_hWnd <> 0) Then
         SetWindowLong m_hWnd, GWL_EXSTYLE, lCStyle
         SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED
      End If
      SetWindowLong UserControl.hWnd, GWL_EXSTYLE, lStyle
      SetWindowPos UserControl.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED
      PropertyChanged "BorderStyle"
   End If
End Property

'Set BackColor
'Public Property Let BackColor(ByVal oNewBackColor As OLE_COLOR)
   'If (oNewBackColor <> m_oBackColor) Then
      'm_oBackColor = oNewBackColor
      'If (m_hWnd <> 0) Then
         'SendMessageLong m_hWnd, SB_SETBKCOLOR, 0, TranslateColor(oNewBackColor)
      'End If
      'PropertyChanged "BackColor"
   'End If
'End Property
'Public Property Get BackColor() As OLE_COLOR
   'BackColor = m_oBackColor
'End Property

'SetForeColor
'Public Property Let ForeColor(ByVal oNewForeColor As OLE_COLOR)
   'If (oNewForeColor <> m_oForeColor) Then
      'm_oForeColor = oNewForeColor
      'If (m_hWnd <> 0) Then
         'SendMessageLong m_hWnd, PBM_SETBARCOLOR, 0, TranslateColor(oNewForeColor)
      'End If
      'PropertyChanged "ForeColor"
   'End If
'End Property

'Public Property Get ForeColor() As OLE_COLOR
   'ForeColor = m_oForeColor
'End Property

Public Property Get Position() As Long
   Position = m_lPosition
End Property
Public Property Let Position(ByVal lPos As Long)
   If (lPos <> m_lPosition) Then
      m_lPosition = lPos
      If (m_hWnd <> 0) Then
         SendMessage m_hWnd, PBM_SETPOS, m_lPosition, 0
      End If
      PropertyChanged "Position"
   End If
End Property
Public Property Get Step() As Long
   Step = m_lStep
End Property
Public Property Let Step(ByVal lStep As Long)
   If (lStep <> m_lStep) Then
      m_lStep = lStep
      If (m_hWnd <> 0) Then
         SendMessage m_hWnd, PBM_SETSTEP, m_lStep, 0
      End If
      PropertyChanged "Step"
   End If
End Property
Public Sub StepIt()
    If (m_hWnd <> 0) Then
        SendMessage m_hWnd, PBM_STEPIT, 0, 0
    Else
        m_lPosition = m_lPosition + m_lStep
    End If
End Sub

Private Sub pSetRange()
Dim tPR As PPBRange
Dim tPA As PPBRange
Dim lR As Long
    If (m_hWnd <> 0) Then
        ' try v4.70 PBM_SETRANGE32:
        SendMessageLong m_hWnd, PBM_SETRANGE32, m_lMin, m_lMax
        
        ' check whether PBM_SETRANGE32 was supported:
        tPA.iHigh = SendMessage(m_hWnd, PBM_GETRANGE, 0, tPR)
        tPA.iLow = SendMessage(m_hWnd, PBM_GETRANGE, 1, tPR)
        If (tPA.iHigh = m_lMax) And (tPA.iLow = m_lMin) Then
            ' ok
        Else
            ' use the original set range message:
            lR = (m_lMin And &HFFFF&)
            CopyMemory VarPtr(lR) + 2, (m_lMax And &HFFFF&), 2
            SendMessage m_hWnd, PBM_SETRANGE, 0, lR
        End If
    End If
End Sub

' Convert Automation color to Windows color
'Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        'Optional hPal As Long = 0) As Long
    'If OleTranslateColor(clr, hPal, TranslateColor) Then
        'TranslateColor = CLR_INVALID
    'End If
'End Function

Private Sub pCreate()
Dim dwStyle As Long
   pDestroy
   InitCommonControls
   dwStyle = WS_VISIBLE Or WS_CHILD
   If (m_eOrientation = epbVertical) Then
      dwStyle = dwStyle Or PBS_VERTICAL
   End If
   If (m_bSmooth) Then
      dwStyle = dwStyle Or PBS_SMOOTH
   End If
   m_hWnd = CreateWindowEX(0, PROGRESS_CLASSA, "", _
              dwStyle, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, _
              UserControl.hWnd, 0&, App.hInstance, 0&)
   If (m_hWnd <> 0) Then
      ' success
      SendMessage m_hWnd, PBM_SETPOS, m_lPosition, 0
   End If
   
End Sub
Private Sub pDestroy()
   If (m_hWnd <> 0) Then
      ShowWindow m_hWnd, SW_HIDE
      SetParent m_hWnd, 0
      DestroyWindow m_hWnd
   End If
End Sub
Private Sub pRecreate()
Dim lPosition As Long
Dim eBorder As EPBBorderStyle
'Dim oBackColor As OLE_COLOR
'Dim oForeColor As OLE_COLOR

   eBorder = BorderStyle
   lPosition = Position
   'oBackColor = BackColor
   'oForeColor = ForeColor
   
   pCreate
   
   pSetRange
   m_lPosition = -1
   Position = m_lPosition
   m_eBorderStyle = -1
   BorderStyle = eBorder
   'm_oBackColor = -1
   'BackColor = oBackColor
   'm_oForeColor = -1
   'ForeColor = oForeColor
   
End Sub

Private Sub UserControl_Initialize()
   m_lMin = 1
   m_lMax = 100
   'm_oForeColor = vbHighlight
   m_lStep = 1
End Sub

Private Sub UserControl_InitProperties()
   Smooth = False
   Orientation = epbHorizontal
   pCreate
   BorderStyle = epbBorderStyleSingle
  ' m_oBackColor = UserControl.Ambient.BackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Smooth = PropBag.ReadProperty("Smooth", False)
   Orientation = PropBag.ReadProperty("Orientation", epbHorizontal)
   pCreate
   m_eBorderStyle = -1
   BorderStyle = PropBag.ReadProperty("BorderStyle", epbBorderStyleSingle)
  ' ForeColor = PropBag.ReadProperty("ForeColor", vbHighlight)
  ' BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
   Min = PropBag.ReadProperty("Min", 0)
   Max = PropBag.ReadProperty("Max", 100)
   Step = PropBag.ReadProperty("Step", 1)
End Sub

Private Sub UserControl_Resize()
   If (m_hWnd <> 0) Then
      MoveWindow m_hWnd, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 1
   End If
End Sub

Private Sub UserControl_Terminate()
   pDestroy
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "BorderStyle", BorderStyle, epbBorderStyleSingle
   PropBag.WriteProperty "Smooth", Smooth, False
   PropBag.WriteProperty "Orientation", Orientation, epbHorizontal
   'PropBag.WriteProperty "ForeColor", m_oForeColor, vbHighlight
   'PropBag.WriteProperty "BackColor", m_oBackColor, vbButtonFace
   PropBag.WriteProperty "Min", Min, 0
   PropBag.WriteProperty "Max", Max, 100
   PropBag.WriteProperty "Step", Step, 1
End Sub


