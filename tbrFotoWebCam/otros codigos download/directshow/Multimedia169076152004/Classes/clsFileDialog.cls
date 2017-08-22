VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FileDialog Class Module by Vesa Piittinen aka Merri
' http://merri.net


Option Explicit


Private Type OPENFILENAME
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
    lpTemplateName As String
End Type


Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


Private FileDialog As OPENFILENAME


'FileDialog class initialize
Private Sub Class_Initialize()
    'init information
    With FileDialog
        .hInstance = App.hInstance
        .hwndOwner = 0
        .lpstrFile = vbNullChar & Space(65534)
        .nMaxFile = 65535
        .lStructSize = Len(FileDialog)
    End With
End Sub


'FileDialog class functions
Public Property Get DefaultExtension() As String
    'return default extension
    DefaultExtension = FileDialog.lpstrDefExt
End Property
Public Property Let DefaultExtension(ByVal Extension As String)
    'set default extension
    FileDialog.lpstrDefExt = Extension
End Property
Public Function FileOpen(ByRef Filename As String) As Boolean
    Dim A As Integer, Temp As String, Temps() As String
    'get filename
    FileOpen = CBool(GetOpenFileName(FileDialog))
    Temps = Split(Trim$(FileDialog.lpstrFile), vbNullChar)
    If IsEmpty(Temps) Then Filename = "": Exit Function
    If UBound(Temps) = 0 Then
        Filename = Temps(0)
    Else
        If Right$(Temps(0), 1) <> "\" Then Temps(0) = Temps(0) & "\"
        For A = 1 To UBound(Temps)
            If Temps(A) <> vbNullString Then Filename = Filename & "|" & Temps(0) & Temps(A)
        Next A
        Filename = Mid$(Filename, 2)
    End If
End Function
Public Function FileSave(ByRef Filename As String) As Boolean
    'get filename
    If Not GetSaveFileName(FileDialog) Then FileSave = False Else Filename = FileDialog.lpstrFile: FileSave = True
End Function
Public Property Let Filter(ByVal CustomFilter As String)
    'replace | with null character
    CustomFilter = Replace$(CustomFilter, "|", vbNullChar)
    'make sure we have a null character in the end
    If Right$(CustomFilter, 1) <> vbNullChar Then CustomFilter = CustomFilter & vbNullChar
    'set filter
    FileDialog.lpstrFilter = CustomFilter
End Property
Public Property Get FilterIndex() As Long
    'return filter index
    FilterIndex = FileDialog.nFilterIndex
End Property
Public Property Let FilterIndex(ByVal Index As Long)
    'set filter index
    FileDialog.nFilterIndex = Index
End Property
Public Property Get Flags() As Long
    'return flags
    Flags = FileDialog.Flags
End Property
Public Property Let Flags(ByVal vFlags As Long)
    'set flags
    FileDialog.Flags = vFlags
End Property
Public Property Get InitialDirectory() As String
    'return initial path
    InitialDirectory = FileDialog.lpstrInitialDir
End Property
Public Property Let InitialDirectory(ByVal InitDir As String)
    'set initial path
    FileDialog.lpstrInitialDir = InitDir
End Property
Public Property Let ObjectOwner(ByVal objOwner As Object)
    'set owner
    FileDialog.hwndOwner = objOwner.hWnd
End Property
Public Property Get WindowTitle() As String
    'return dialog title
    WindowTitle = FileDialog.lpstrTitle
End Property
Public Property Let WindowTitle(ByVal Title As String)
    'set dialog title
    FileDialog.lpstrTitle = Title
End Property
