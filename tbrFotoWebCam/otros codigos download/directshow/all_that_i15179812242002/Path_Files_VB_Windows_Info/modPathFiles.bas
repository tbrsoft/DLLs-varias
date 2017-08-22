Attribute VB_Name = "modPathFiles"
Option Explicit

Public Const MAX_PATH = 260

Private Enum FILE_ATTRIBUTE
  FILE_ATTRIBUTE_DIRECTORY = &H10
  FILE_ATTRIBUTE_ARCHIVE = &H20
  FILE_ATTRIBUTE_NORMAL = &H80
  FILE_ATTRIBUTE_READONLY = &H1
  FILE_ATTRIBUTE_HIDDEN = &H2
  FILE_ATTRIBUTE_SYSTEM = &H4
  FILE_ATTRIBUTE_COMPRESSED = &H800
  FILE_ATTRIBUTE_ENCRYPTED = &H40
  FILE_ATTRIBUTE_TEMPORARY = &H100
  FILE_ATTRIBUTE_OFFLINE = &H1000
  FILE_ATTRIBUTE_SPARSE_FILE = &H200
  FILE_ATTRIBUTE_REPARSE_POINT = &H400
  FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
End Enum

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME       ' DayOfWeek :
    wYear As Integer         ' ------------
    wMonth As Integer        ' Dimanche = 0
    wDayOfWeek As Integer    ' Lundi    = 1
    wDay As Integer          ' Mardi    = 2
    wHour As Integer         ' Mercredi = 3
    wMinute As Integer       ' Jeudi    = 4
    wSecond As Integer       ' Vendredi = 5
    wMilliseconds As Integer ' Samedi   = 6
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As FILE_ATTRIBUTE
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Type Search_File_Type
    dwFileAttributes As FILE_ATTRIBUTE
    nFileSize As Currency
    cPath As Variant
    cFileName As Variant
    cPathAndFileName As Variant
    stCreationTime As SYSTEMTIME
    stLastAccessTime As SYSTEMTIME
    stLastWriteTime As SYSTEMTIME
End Type

Public Enum DriveTypeVar
  DRIVE_ERROR = -1
  DRIVE_UNKNOWN = 0
  DRIVE_ABSENT = 1
  DRIVE_REMOVABLE = 2
  DRIVE_FIXED = 3
  DRIVE_REMOTE = 4
  DRIVE_CDROM = 5
  DRIVE_RAMDISK = 6
End Enum

Private Type BrowseInfo
  hWndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SetCurrentDirectory Lib "kernel32.dll" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SearchTreeForFile Lib "imagehlp.dll" (ByVal RootPath As String, ByVal InputPathName As String, ByVal OutputPathBuffer As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectoryEmpty Lib "shlwapi.dll" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal lBuffer As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private m_CurrentDirectory As String
Private FileArray() As Search_File_Type
Private TotalFilesFound As Long

Public Function AppPath(ByVal zPath As String) As String
  If Right$(zPath, 1) = "\" Then AppPath = zPath Else AppPath = zPath & "\"
End Function

Private Function GetDT(ByVal var1 As String) As String
  Select Case GetDriveType(var1)
    Case DRIVE_UNKNOWN
      GetDT = "UNKNOWN"
    Case DRIVE_ABSENT
      GetDT = "ABSENT"
    Case DRIVE_REMOVABLE
      GetDT = "REMOVABLE"
    Case DRIVE_FIXED
      GetDT = "FIXED"
    Case DRIVE_REMOTE
      GetDT = "REMOTE"
    Case DRIVE_CDROM
      GetDT = "CDROM"
    Case DRIVE_RAMDISK
      GetDT = "RAMDISK"
    Case Else
      GetDT = "ERROR"
  End Select
End Function

Public Function DriveTypeS(ByVal zDrive As String) As String
  Dim var1 As String
  If ((Len(zDrive) = 2) And (Right$(zDrive, 1) = ":")) Or (Len(zDrive) = 1) Then
    var1 = Left$(zDrive, 1)
    Select Case Asc(var1)
      Case 65 To 90
        DriveTypeS = GetDT(var1 & ":")
      Case 97 To 122
        DriveTypeS = GetDT(var1 & ":")
      Case Else
        MsgBox "Use " & Chr$(34) & "DriveTypeS c" & Chr$(34) & " or " & Chr$(34) & "DriveTypeS c:" & Chr$(34), vbExclamation, "Super"
        DriveTypeS = "ERROR"
    End Select
  Else
    MsgBox "Use " & Chr$(34) & "DriveTypeS c" & Chr$(34) & " or " & Chr$(34) & "DriveTypeS c:" & Chr$(34), vbExclamation, "Super"
    DriveTypeS = "ERROR"
  End If
End Function

Public Function DriveType(ByVal zDrive As String) As DriveTypeVar
  Dim var1 As String
  If ((Len(zDrive) = 2) And (Right$(zDrive, 1) = ":")) Or (Len(zDrive) = 1) Then
    var1 = Left$(zDrive, 1)
    Select Case Asc(var1)
      Case 65 To 90
        DriveType = GetDriveType(var1 & ":")
      Case 97 To 122
        DriveType = GetDriveType(var1 & ":")
      Case Else
        MsgBox "Use " & Chr$(34) & "DriveType c" & Chr$(34) & " or " & Chr$(34) & "DriveType c:" & Chr$(34), vbExclamation, "Super"
        DriveType = DRIVE_ERROR
    End Select
  Else
    MsgBox "Use " & Chr$(34) & "DriveType c" & Chr$(34) & " or " & Chr$(34) & "DriveType c:" & Chr$(34), vbExclamation, "Super"
    DriveType = DRIVE_ERROR
  End If
End Function

Public Function FileExist(ByVal strPath As String) As Boolean
  On Local Error GoTo ErrFile
  Open strPath For Input Access Read As #1
  Close #1
  FileExist = True
  Exit Function
ErrFile:
  FileExist = False
End Function

Public Function Filexist(ByVal strPath As String) As Boolean
  Filexist = FileExist(strPath)
End Function

Public Function DirExist(ByVal zPath As String) As Boolean
'  On Local Error GoTo ErrDir
  DirExist = PathIsDirectory(zPath)
'  Dim qwe As String
'  qwe = CurDir
'  ChDir zPath
'  ChDir qwe
'  DirExist = True
'  Exit Function
'ErrDir:
'  DirExist = False
End Function

Public Function SetCurDir(ByVal zPath As String) As Boolean
  If SetCurrentDirectory(zPath) <> 0 Then
    SetCurDir = True
  Else
    SetCurDir = False
  End If
End Function

Public Function FreeSpace(ByVal zDrive As String) As Currency
  Dim var1 As Currency, var2 As Currency, var3 As Currency, var4 As String
  If ((Len(zDrive) = 2) And (Right$(zDrive, 1) = ":")) Or (Len(zDrive) = 1) Then
    var4 = Left$(zDrive, 1)
    Select Case Asc(var4)
      Case 65 To 90
        GetDiskFreeSpaceEx var4 & ":", var1, var2, var3
      Case 97 To 122
        GetDiskFreeSpaceEx var4 & ":", var1, var2, var3
      Case Else
        MsgBox "Use " & Chr$(34) & "FreeSpace c" & Chr$(34) & " or " & Chr$(34) & "FreeSpace c:" & Chr$(34), vbExclamation, "Super"
        FreeSpace = -1
        Exit Function
    End Select
  Else
    MsgBox "Use " & Chr$(34) & "FreeSpace c" & Chr$(34) & " or " & Chr$(34) & "FreeSpace c:" & Chr$(34), vbExclamation, "Super"
    FreeSpace = -1
    Exit Function
  End If
  If var1 = 0 And var2 = 0 And var3 = 0 Then
    FreeSpace = -1
  Else
    FreeSpace = var1 * 10000
  End If
End Function

Public Function FileOrDirExist(ByVal zPath As String) As Boolean
  FileOrDirExist = PathFileExists(zPath)
End Function

Public Function TreeFind(ByVal zPath As String, ByVal zFile As String) As Variant
  Dim VarTemp As String
  VarTemp = String(MAX_PATH, 0)
  If SearchTreeForFile(zPath, zFile, VarTemp) <> 0 Then
    TreeFind = Trim2(VarTemp)
  Else
    TreeFind = -1
  End If
End Function

Private Function SearchFiles(ByVal zPath As String, ByVal zFiles As String, Optional ByVal SubDirs As Boolean = True, Optional ByRef NumberFound As Long = -1, Optional NewSearch As Boolean = True) As Search_File_Type()
  Dim zPathStr As String, DirCount As Long, FileCount As Long, isOK As Boolean
  Dim RetVal As Long, TempSearch() As WIN32_FIND_DATA, DDir() As String, t As Long
  If NewSearch = True Then TotalFilesFound = 0
  If zFiles = vbNullString Or zFiles = "" Then zFiles = "*.*"
  If Right$(zPath, 1) = "\" Then
    zPathStr = zPath
  Else
    zPathStr = zPath & "\"
  End If
  DirCount = 0
  isOK = True
  ReDim TempSearch(1 To 1)
  RetVal = FindFirstFile(zPathStr & "*.*", TempSearch(1))
  If RetVal <> -1 Then
  Do While isOK
    DoEvents
    If (FILE_ATTRIBUTE_DIRECTORY And TempSearch(1).dwFileAttributes) = FILE_ATTRIBUTE_DIRECTORY Then
      If Trim2(TempSearch(1).cFileName) <> "." And Trim2(TempSearch(1).cFileName) <> ".." Then
        DirCount = DirCount + 1
        ReDim Preserve DDir(DirCount)
        DDir(DirCount) = Trim2(TempSearch(1).cFileName)
      End If
    End If
    ReDim TempSearch(1 To 1)
    isOK = FindNextFile(RetVal, TempSearch(1)) <> 0
  Loop
  End If
  FindClose RetVal
  FileCount = 0
  isOK = True
  ReDim TempSearch(1 To 1)
  RetVal = FindFirstFile(zPathStr & zFiles, TempSearch(1))
  If RetVal <> -1 Then
  Do While isOK
    DoEvents
    If Trim2(TempSearch(1).cFileName) <> "." And Trim2(TempSearch(1).cFileName) <> ".." Then
      FileCount = FileCount + 1
      ReDim Preserve FileArray(TotalFilesFound + 1)
      FileArray(TotalFilesFound + 1).cPath = zPathStr
      FileArray(TotalFilesFound + 1).cFileName = Trim2(TempSearch(1).cFileName)
      FileArray(TotalFilesFound + 1).cPathAndFileName = zPathStr & Trim2(TempSearch(1).cFileName)
      FileArray(TotalFilesFound + 1).dwFileAttributes = TempSearch(1).dwFileAttributes
      FileArray(TotalFilesFound + 1).nFileSize = (4294967296@ * TempSearch(1).nFileSizeHigh) + TempSearch(1).nFileSizeLow
      FileTimeToSystemTime TempSearch(1).ftCreationTime, FileArray(TotalFilesFound + 1).stCreationTime
      FileTimeToSystemTime TempSearch(1).ftLastWriteTime, FileArray(TotalFilesFound + 1).stLastWriteTime
      FileTimeToSystemTime TempSearch(1).ftLastAccessTime, FileArray(TotalFilesFound + 1).stLastAccessTime
      TotalFilesFound = TotalFilesFound + 1
      If NumberFound <> -1 Then NumberFound = TotalFilesFound
    End If
    ReDim TempSearch(1 To 1)
    isOK = FindNextFile(RetVal, TempSearch(1)) <> 0
  Loop
  End If
  FindClose RetVal
  If SubDirs = True Then
    If NumberFound = -1 Then
      For t = 1 To DirCount
        DoEvents
        SearchFiles zPathStr & DDir(t), zFiles, True, , False
      Next t
    Else
      For t = 1 To DirCount
        DoEvents
        SearchFiles zPathStr & DDir(t), zFiles, True, NumberFound, False
      Next t
    End If
  End If
  If NumberFound <> -1 Then NumberFound = TotalFilesFound
  SearchFiles = FileArray
End Function

Public Function StripPath(ByVal zPathAndFile As String) As String
  Dim t As Long
  For t = Len(zPathAndFile) To 1 Step -1
    If Mid$(zPathAndFile, t, 1) = "\" Or Mid$(zPathAndFile, t, 1) = ":" Then Exit For
  Next t
  StripPath = Mid$(zPathAndFile, t + 1)
End Function

Public Function StripFile(ByVal zPathAndFile As String) As String
  Dim t As Long
  For t = Len(zPathAndFile) To 1 Step -1
    If Mid$(zPathAndFile, t, 1) = "\" Or Mid$(zPathAndFile, t, 1) = ":" Then Exit For
  Next t
  StripFile = Left$(zPathAndFile, t)
End Function

Public Function IsDirEmpty(ByVal zPath As String) As Boolean
  IsDirEmpty = PathIsDirectoryEmpty(zPath)
End Function

Public Function GetShortPath(ByVal zPathAndFile As String) As String
  Dim StrLen As Long, ShortPath As String
  ShortPath = String$(MAX_PATH, 0)
  StrLen = GetShortPathName(zPathAndFile, ShortPath, MAX_PATH)
  GetShortPath = Left$(ShortPath, StrLen)
End Function

Public Function GetLongPath(ByVal zPathAndFile As String) As String
  Dim StrLen As Long, LongPath As String
  LongPath = String$(MAX_PATH, 0)
  StrLen = GetLongPathName(zPathAndFile, LongPath, MAX_PATH)
  GetLongPath = Left$(LongPath, StrLen)
End Function

Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function

Private Function BrowseCallbackProc(ByVal HWND As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
  On Local Error Resume Next
  Dim lpIDList As Long
  Dim ret As Long
  Dim sBuffer As String
  Select Case uMsg
    Case BFFM_INITIALIZED
      SendMessage HWND, BFFM_SETSELECTION, 1, m_CurrentDirectory
    Case BFFM_SELCHANGED
      sBuffer = Space(MAX_PATH)
      ret = SHGetPathFromIDList(lp, sBuffer)
      If ret = 1 Then
        SendMessage HWND, BFFM_SETSTATUSTEXT, 0, sBuffer
      End If
  End Select
  BrowseCallbackProc = 0
End Function

Public Function BrowseForFolder(ByVal Title As String, ByVal StartDir As String, Optional owner As Form = Nothing, Optional IncludeFiles As Boolean = False) As String
  Const BIF_STATUSTEXT = &H4
  Const BIF_RETURNONLYFSDIRS = &H1
  Const BIF_BROWSEINCLUDEFILES = &H4000
  Dim lpIDList As Long, sBuffer As String, tBrowseInfo As BrowseInfo
  If Len(StartDir) > 0 Then m_CurrentDirectory = StartDir & vbNullChar
  If Len(Title) > 0 Then
    tBrowseInfo.lpszTitle = lstrcat(Title, "")
  Else
    tBrowseInfo.lpszTitle = lstrcat("Select A Directory", "")
  End If
  If Not (owner Is Nothing) Then tBrowseInfo.hWndOwner = owner.HWND
  If IncludeFiles = True Then
    tBrowseInfo.ulFlags = BIF_STATUSTEXT + BIF_RETURNONLYFSDIRS + BIF_BROWSEINCLUDEFILES
  Else
    tBrowseInfo.ulFlags = BIF_STATUSTEXT + BIF_RETURNONLYFSDIRS
  End If
  tBrowseInfo.lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    CoTaskMemFree lpIDList
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = sBuffer
  Else
    BrowseForFolder = ""
  End If
End Function
