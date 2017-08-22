Attribute VB_Name = "mdGlobals"
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

'=========================================================================
' Constant and variables
'=========================================================================

Public Const LIB_NAME               As String = "VbAmpPlayer"
Public Const PROGID_DOCUMENT        As String = LIB_NAME & ".docPlayer"
Public Const STR_APP_NAME           As String = "VbAmp Player"
Public Const MASK_COLOR             As Long = &HFF00FF

Public g_oVbIde                 As VBIDE.VBE
Public g_oAddinWindow           As VBIDE.Window

'=========================================================================
' Functions
'=========================================================================

