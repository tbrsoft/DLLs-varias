Attribute VB_Name = "modCommon"
Option Explicit
'==========================================================================
'  This is a part of Banasoft AVPhone controls
'  To get the last version of the control, please visit:
'
'  http://www.banasoft.net/AVPhone.htm
'
'  THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
'  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
'  PURPOSE.
'
'  Copyright (c) - 2002  Banasoft.  All Rights Reserved.
'
'==========================================================================

Public Sub ShowErr()
    MsgBox Err.Description, vbCritical
End Sub


Public Sub ShowCode(AppPath As String, ParamArray FileList() As Variant)

    Dim sp As String
    sp = App.Path
    If Right$(sp, 1) <> "\" Then sp = sp & "\"
    
    
    Dim s As String
    
    Dim v As Variant
    For Each v In FileList
        If Len(s) <= 0 Then
            s = """" & sp & v & """"
        Else
            s = """" & sp & v & """ " & s
        End If
    Next
    
    Shell sp & AppPath & "..\..\..\srcview.exe " & s, vbNormalFocus
End Sub
