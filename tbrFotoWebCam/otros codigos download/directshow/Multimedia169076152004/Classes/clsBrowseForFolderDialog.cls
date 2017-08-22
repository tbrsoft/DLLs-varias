VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrowseForFolderDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' BrowseForFolderDialog Class Module by Vesa Piittinen aka Merri
' http://merri.net


Option Explicit


Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Const MAX_PATH = 260

Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2


Private Type BROWSEINFO
    howner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    lImage As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long


Private bi As BROWSEINFO, pidl As Long, Owner As Long
Private CreateFolder As Boolean, DefaultFolder As String, DialogMessage As String


'BrowseForFolderDialog class initialize
Private Sub Class_Initialize()
    DefaultFolder = "C:\"
    DialogMessage = "Select a folder:"
End Sub


'BrowseForFolderDialog class functions
Public Function FolderDialog(ByRef Path As String) As Boolean
    'do not show new folder button
    CreateFolder = False
    'get path
    Path = GetFolderPath
    'success if not empty string
    FolderDialog = Path <> vbNullString
End Function
Public Property Get InitialDirectory() As String
    'return initial path
    InitialDirectory = DefaultFolder
End Property
Public Property Let InitialDirectory(ByVal InitDir As String)
    'make sure default folder ends in character \
    If Right$(InitDir, 1) <> "\" Then InitDir = InitDir & "\"
    'set initial path
    DefaultFolder = InitDir
End Property
Public Property Let ObjectOwner(objOwner As Object)
    'set owner
    Owner = objOwner.hWnd
End Property
Public Function NewFolderDialog(ByRef Path As String) As Boolean
    'show new folder button
    CreateFolder = True
    'get path
    Path = GetFolderPath
    'success if not empty string
    NewFolderDialog = Path <> vbNullString
End Function
Public Property Get WindowMessage() As String
    'return dialog title
    WindowMessage = DialogMessage
End Property
Public Property Let WindowMessage(ByVal Message As String)
    'set dialog title
    DialogMessage = Message
End Property


'BrowseForFolderDialog class additional functions
Private Function GetFolderPath() As String
    Dim lpSelPath As Long, sPath As String * MAX_PATH, pidl As Long, iNull As Integer, strFolderPath As String
    sPath = DefaultFolder
    'reserve memory
    lpSelPath = LocalAlloc(LPTR, Len(sPath) + 1)
    CopyMemory ByVal lpSelPath, ByVal sPath, Len(sPath) + 1
    With bi
        .howner = Owner
        .pidlRoot = 0
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        .lParam = lpSelPath
        .lpszTitle = DialogMessage & vbNullChar
        If CreateFolder Then .ulFlags = BIF_NEWDIALOGSTYLE Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN
    End With
    pidl = SHBrowseForFolder(bi)
    If pidl Then
        'get actual string
        If SHGetPathFromIDList(pidl, sPath) Then
            'trim path
            strFolderPath = Trim$(sPath)
            'check for null character
            iNull = InStr(strFolderPath, vbNullChar)
            If iNull Then strFolderPath = Left$(strFolderPath, iNull - 1)
        End If
        CoTaskMemFree pidl
    End If
    LocalFree lpSelPath
    'return path
    GetFolderPath = strFolderPath
End Function
