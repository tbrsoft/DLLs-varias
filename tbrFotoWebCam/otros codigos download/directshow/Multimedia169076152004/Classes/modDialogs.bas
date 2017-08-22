Attribute VB_Name = "modDialogs"
' Dialogs Module by Vesa Piittinen aka Merri
' http://merri.net
'
' additional and required module for clsBrowseForFolderDialog and clsFileDialog


Option Explicit


'FileDialog Class Module Constants
Public Const clsFDReadOnly = &H1                    'Checks Read-Only check box for Open and Save As dialog boxes.
Public Const clsFDOverwritePrompt = &H2             'Causes the Save As dialog box to generate a message box if the selected file already exists.
Public Const clsFDHideReadOnly = &H4                'Hides the Read-Only check box.
Public Const clsFDNoChangeDir = &H8                 'Sets the current directory to what it was when the dialog box was invoked.
Public Const clsFDHelpButton = &H10                 'Causes the dialog box to display the Help button.
Public Const clsFDNoValidate = &H100                'Allows invalid characters in the returned filename.
Public Const clsFDAllowMultiselect = &H200          'Allows the File Name list box to have multiple selections.
Public Const clsFDExtensionDifferent = &H400        'The extension of the returned filename is different from the extension set by the DefaultExt property.
Public Const clsFDPathMustExist = &H800             'User can enter only valid path names.
Public Const clsFDFileMustExist = &H1000            'User can enter only names of existing files.
Public Const clsFDCreatePrompt = &H2000             'Sets the dialog box to ask if the user wants to create a file that doesn't currently exist.
Public Const clsFDShareAware = &H4000               'Sharing violation errors will be ignored.
Public Const clsFDNoReadOnlyReturn = &H8000         'The returned file doesn't have the Read-Only attribute set and won't be in a write-protected directory.
Public Const clsFDExplorer = &H80000                'Use the Explorer-like Open A File dialog box template.
Public Const clsFDNoDereferenceLinks = &H100000     'Do not dereference shortcuts (shell links).  By default, choosing a shortcut causes it to be dereferenced by the shell.
Public Const clsFDLongNames = &H200000              'Use Long filenames.


'BrowseForFolderDialog Constants
Private Const WM_USER = &H400

Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELECTIONCHANGED = 2

Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)


'BrowseForFolderDialog API Declarations
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


'BrowseForFolderDialog Class Module Additional Help Functions
Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End Select
End Function
Public Function FARPROC(ByVal pfn As Long) As Long
    'dummy procedure returning value of AddressOf
    FARPROC = pfn
End Function
