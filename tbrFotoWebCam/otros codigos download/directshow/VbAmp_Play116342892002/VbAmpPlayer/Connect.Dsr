VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   8784
   ClientLeft      =   1740
   ClientTop       =   1548
   ClientWidth     =   9876
   _ExtentX        =   17420
   _ExtentY        =   15494
   _Version        =   393216
   Description     =   "VbAmp Player is an all purpose dockable player capable of playing AVI, MPEG and MP3 files."
   DisplayName     =   " VbAmp Player"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
'   You are free to use this source as long as this copyright message
'     appears on your program's "About" dialog:
'
'   VbAmp Player Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqw@myrealbox.com)
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "Connect"

'=========================================================================
' Constand and member variables
'=========================================================================

Private Const CAP_MSG                   As String = "Add-in connect"
Private Const STR_MAIN_MENU_CAPTION     As String = "Add-Ins"
Private Const STR_STANDARD_CAPTION      As String = "Standard"
Private Const STR_GUID_POSITION         As String = "AE417875-8082-449f-A66F-7992D1F9CC63"
Private Const RID_BMP_MENU              As Long = 109

Private m_oMenu                     As Office.CommandBarButton
Private WithEvents m_oMenuEvents    As VBIDE.CommandBarEvents
Attribute m_oMenuEvents.VB_VarHelpID = -1
Private m_oToolbar                  As Office.CommandBarButton
Private WithEvents m_oToolbarEvents As VBIDE.CommandBarEvents
Attribute m_oToolbarEvents.VB_VarHelpID = -1
Private m_docPlayer                 As docPlayer

'=========================================================================
' Error handling
'=========================================================================

'Private Sub RaiseError(sFunc As String)
'    PushError sFunc, MODULE_NAME
'    PopRaiseError
'End Sub

Private Function ShowError(sFunc As String) As VbMsgBoxResult
    PushError sFunc, MODULE_NAME
    ShowError = PopShowError(CAP_MSG)
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Const FUNC_NAME     As String = "AddinInstance_OnConnection"

    On Error GoTo EH
    '--- store global reference to VB IDE
    Set g_oVbIde = Application
    '--- create (transparent) menu icon and store in clipboard
    With New cMemDC
        .Init 16, 16
        .Cls vbButtonFace
        .PaintPicture LoadResPicture(RID_BMP_MENU, vbResBitmap), clrMask:=MASK_COLOR
        Clipboard.Clear
        Clipboard.SetData .Image
    End With
    '--- add sub-menu item with icon
    Set m_oMenu = g_oVbIde.CommandBars(STR_MAIN_MENU_CAPTION).Controls.Add(msoControlButton)
    m_oMenu.Caption = STR_APP_NAME
    m_oMenu.PasteFace
    '--- add toolbar button with icon
    With g_oVbIde.CommandBars(STR_STANDARD_CAPTION).Controls
        If .Item(.Count).Type <> msoControlButton Then
            Set m_oToolbar = .Add(msoControlButton, , , .Count)
        Else
            Set m_oToolbar = .Add(msoControlButton)
        End If
    End With
    m_oToolbar.Caption = STR_APP_NAME
    m_oToolbar.PasteFace
    '--- sink events
    Set m_oMenuEvents = g_oVbIde.Events.CommandBarEvents(m_oMenu)
    Set m_oToolbarEvents = g_oVbIde.Events.CommandBarEvents(m_oToolbar)
    '--- create toolwindow
    Set g_oAddinWindow = g_oVbIde.Windows.CreateToolWindow( _
                AddInInst, _
                PROGID_DOCUMENT, _
                STR_APP_NAME, _
                STR_GUID_POSITION, _
                m_docPlayer)
    '--- read settings
    m_docPlayer.ReadSettings
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Const FUNC_NAME     As String = "AddinInstance_OnDisconnection"

    On Error GoTo EH
    '--- delete the command bar entry
    If Not m_oMenu Is Nothing Then
        m_oMenu.Delete
        Set m_oMenu = Nothing
    End If
    Set m_oMenuEvents = Nothing
    '--- delete toolbar entry
    If Not m_oToolbar Is Nothing Then
        m_oToolbar.Delete
        Set m_oToolbar = Nothing
    End If
    Set m_oToolbarEvents = Nothing
    '--- persist settings and clear reference
    If Not m_docPlayer Is Nothing Then
        m_docPlayer.SaveSettings
        Set m_docPlayer = Nothing
    End If
    '--- cleanup global references
    Set g_oAddinWindow = Nothing
    Set g_oVbIde = Nothing
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub m_oMenuEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Const FUNC_NAME     As String = "m_oMenuEvents_Click"

    On Error GoTo EH
    If CommandBarControl.Caption = STR_APP_NAME Then
        g_oAddinWindow.Visible = True
    End If
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub m_oToolbarEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Const FUNC_NAME     As String = "m_oToolbarEvents_Click"

    On Error GoTo EH
    If CommandBarControl.Caption = STR_APP_NAME Then
        g_oAddinWindow.Visible = True
    End If
    Exit Sub
EH:
    Select Case ShowError(FUNC_NAME)
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

