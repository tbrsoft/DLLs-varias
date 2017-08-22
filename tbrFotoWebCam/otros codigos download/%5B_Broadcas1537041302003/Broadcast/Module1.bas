Attribute VB_Name = "Module1"
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

'for common dialog functions

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long

Public Enum EOpenFile
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
End Enum

'get a open file path
Public Function OpenFileDlg(ByVal hwnd As Long, ByVal Index As Long) As String

    Dim of As OPENFILENAME
    With of
        .lStructSize = Len(of)
        .Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
        .hwndOwner = hwnd
        .lpstrDefExt = "avi"
        .lpstrFilter = "AVI Files (*.avi)" & vbNullChar & "*.avi" & vbNullChar & "Wave Files (*.wav)" & vbNullChar & "*.wav" & vbNullChar & "Bitmap Files (*.bmp)" & vbNullChar & "*.bmp" & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        .nFilterIndex = Index
        .nMaxFile = 256
        
        Dim s As String
        s = String$(256, 0)
        .lpstrFile = s
        
        If GetOpenFileName(of) Then
        
            Dim l As Long
            l = InStr(.lpstrFile, vbNullChar)
            If l > 0 Then
                OpenFileDlg = Left$(.lpstrFile, l - 1)
            Else
                OpenFileDlg = .lpstrFile
            End If
            
        Else
            Err.Raise 32755
        End If
    End With
End Function


'get a save file path
Public Function SaveASFileDlg(ByVal hwnd As Long, ByVal Index As Long) As String
            
    Dim of As OPENFILENAME
    With of
        .lStructSize = Len(of)
        .Flags = (OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY)
        .hwndOwner = hwnd
        .lpstrDefExt = "avi"
        .lpstrFilter = "AVI Files (*.avi)" & vbNullChar & "*.avi" & vbNullChar & "Wave Files (*.wav)" & vbNullChar & "*.wav" & vbNullChar & "Bitmap Files (*.bmp)" & vbNullChar & "*.bmp" & vbNullChar & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        .nFilterIndex = Index
        .nMaxFile = 256
    
        Dim s As String
        s = String$(256, 0)
        .lpstrFile = s
    
        If GetSaveFileName(of) Then
            Dim l As Long
            l = InStr(.lpstrFile, vbNullChar)
            If l > 0 Then
                SaveASFileDlg = Left$(.lpstrFile, l - 1)
            Else
                SaveASFileDlg = .lpstrFile
            End If
        Else
            Err.Raise 32755
        End If
    End With
End Function
