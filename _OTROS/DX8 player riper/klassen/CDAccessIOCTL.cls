VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDAccessIOCTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Low Level CD access through IOCTLs
'
' + you don't need write access to the drive
' - only NT/2K/XP, no inquiry!

Implements ICDAccess

Private Type IOCTLDrive
    char                As String
    handle              As hFile
    vendor              As String
    product             As String
    revision            As String
End Type

Private Const CDROM_SECTOR_SIZE As Long = 2352
Private Const CDROM_DATA_TO_SEC As Long = 2048

Private udtDrives()     As IOCTLDrive

Private lngDriveCount   As Long
Private lngSelected     As Long

Private blnInitialized  As Boolean

Private Sub GetDeviceName( _
    ByVal index As Long _
)

    ' no INQUIRY IOCTL...
    ' seems like it only exists for Win CE Net 4

    With udtDrives(index)
        .product = "CD-ROM drive"
        .vendor = "Generic"
        .revision = "v1"
    End With
End Sub

Private Sub GetDriveList()
    Dim i       As Integer
    Dim hDrive  As hFile

    lngDriveCount = 0

    For i = 2 To 26
        If GetDriveType(Chr$(i + 64) & ":") = DTYPE_CDROM Then
            hDrive = GetDriveHandle(Chr$(i + 64), GENERIC_READ)
            If hDrive.handle <> INVALID_HANDLE Then
                ReDim Preserve udtDrives(lngDriveCount) As IOCTLDrive

                udtDrives(lngDriveCount).handle = hDrive
                udtDrives(lngDriveCount).char = Chr$(i + 64)
                GetDeviceName lngDriveCount

                lngDriveCount = lngDriveCount + 1
            End If
        End If
    Next
End Sub

Private Sub Class_Initialize()
    modDevIoCtl.InitIOCTLs
    lngSelected = -1
End Sub

Private Sub Class_Terminate()
    ICDAccess_DeInitMe
End Sub

Private Function ICDAccess_CloseCD() As STREAM_ERR
    Dim dwRet       As Long
    Dim dwBytesRet  As Long

    If Not blnInitialized Then Exit Function
    If lngSelected = -1 Then Exit Function

    dwRet = DeviceIoControl(udtDrives(lngSelected).handle.handle, _
                            IOCTL_DISK_LOAD_MEDIA, _
                            ByVal 0&, 0, _
                            ByVal 0&, 0, _
                            dwBytesRet, _
                            ByVal 0&)

    If dwRet = 1 Then
        ICDAccess_CloseCD = STREAM_OK
    Else
        ICDAccess_CloseCD = STREAM_NOT_SUPPORTED
    End If
End Function

Private Function ICDAccess_DeInitMe() As STREAM_ERR
    Dim i   As Integer

    If Not blnInitialized Then
        ICDAccess_DeInitMe = STREAM_OK
        Exit Function
    End If

    For i = 0 To lngDriveCount - 1
        FileClose udtDrives(i).handle
    Next
    lngDriveCount = 0

    ICDAccess_DeInitMe = STREAM_OK
End Function

Private Function ICDAccess_DeviceChar( _
    ByVal index As Long _
) As String

    ICDAccess_DeviceChar = udtDrives(index).char
End Function

Private Function ICDAccess_DeviceCount() As Long
    ICDAccess_DeviceCount = lngDriveCount
End Function

Private Function ICDAccess_DeviceName( _
    ByVal index As Long _
) As String

    With udtDrives(index)
        ICDAccess_DeviceName = .vendor & " " & .product & " " & .revision
    End With
End Function

Private Function ICDAccess_GetTOC( _
    refTOC As CDTOC _
) As STREAM_ERR

    Dim dwRet       As Long
    Dim dwBytesRet  As Long
    Dim i           As Long
    Dim LBA         As Long
    Dim udtTOC      As CDROM_TOC
    Dim clsTOC      As CDTOC

    Set clsTOC = New CDTOC

    If Not blnInitialized Then Exit Function
    If lngSelected = -1 Then Exit Function

    dwRet = DeviceIoControl(udtDrives(lngSelected).handle.handle, _
                            IOCTL_CDROM_READ_TOC, _
                            ByVal 0&, 0, _
                            udtTOC, Len(udtTOC), _
                            dwBytesRet, ByVal 0&)

    If dwRet = 0 Then
        ICDAccess_GetTOC = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    If clsTOC.TrackCount = 0 Then
        ICDAccess_GetTOC = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    If udtTOC.LastTrack = 0 Then
        ICDAccess_GetTOC = STREAM_NOT_SUPPORTED
        Exit Function
    End If

    clsTOC.TrackCount = udtTOC.LastTrack + 1

    For i = 1 To udtTOC.LastTrack + 1
        ' MSF -> LBA
        LBA = MSF2LBA(VarToMSF(udtTOC.TrackData(i - 1).addr))
        With udtTOC.TrackData(i - 1)
            .addr(0) = SHR(LBA, 24) And &HFF
            .addr(1) = SHR(LBA, 16) And &HFF
            .addr(2) = SHR(LBA, 8) And &HFF
            .addr(3) = LBA And &HFF
        End With
        clsTOC.SetTrackData udtTOC.TrackData(i - 1)
    Next

    Set refTOC = clsTOC

    ICDAccess_GetTOC = STREAM_OK
End Function

Private Function ICDAccess_InitMe() As STREAM_ERR
    ICDAccess_InitMe = STREAM_NOT_SUPPORTED

    ' only NT/2k/XP
    If GetWindowsVersion.dwPlatformId <> VER_PLATFORM_WIN32_NT Then
        Exit Function
    End If

    GetDriveList
    If lngDriveCount > 0 Then
        blnInitialized = True
        ICDAccess_InitMe = STREAM_OK
    End If
End Function

Private Function ICDAccess_IsDiscPresent() As Boolean
    Dim udtGeometry As DISK_GEOMETRY
    Dim dwRet       As Long
    Dim dwBytesRet  As Long

    If Not blnInitialized Then Exit Function
    If lngSelected = -1 Then Exit Function

    dwRet = DeviceIoControl(udtDrives(lngSelected).handle.handle, _
                            IOCTL_DISK_GET_DRIVE_GEOMETRY, _
                            ByVal 0&, 0, _
                            udtGeometry, Len(udtGeometry), _
                            dwBytesRet, ByVal 0&)

    ICDAccess_IsDiscPresent = dwRet = 1
End Function

Private Function ICDAccess_OpenCD() As STREAM_ERR
    Dim dwRet       As Long
    Dim dwBytesRet  As Long

    If Not blnInitialized Then Exit Function
    If lngSelected = -1 Then Exit Function

    dwRet = DeviceIoControl(udtDrives(lngSelected).handle.handle, _
                            IOCTL_DISK_EJECT_MEDIA, _
                            ByVal 0&, 0, _
                            ByVal 0&, 0, _
                            dwBytesRet, _
                            ByVal 0&)

    If dwRet = 1 Then
        ICDAccess_OpenCD = STREAM_OK
    Else
        ICDAccess_OpenCD = STREAM_NOT_SUPPORTED
    End If
End Function

' http://www.insidepro.com/kk/020/020r.shtml
Private Function ICDAccess_ReadSectorsRAW( _
    ByVal LBA As Long, _
    ByVal sectors As Long, _
    ByVal ptr As Long _
) As STREAM_ERR

    Dim dwRet       As Long
    Dim dwBytesRet  As Long
    Dim udtRead     As RAW_READ_INFO

    If Not blnInitialized Then Exit Function
    If lngSelected = -1 Then Exit Function

    With udtRead
        .TrackMode = TRACK_MODE_CDDA
        .SectorCount = sectors
        .DiskOffset = Dbl2LargeInt(LBA * CDROM_DATA_TO_SEC)
    End With

    dwRet = DeviceIoControl(udtDrives(lngSelected).handle.handle, _
                            IOCTL_CDROM_RAW_READ, _
                            udtRead, Len(udtRead), _
                            ByVal ptr, CDROM_SECTOR_SIZE * sectors, _
                            dwBytesRet, ByVal 0&)

    If dwRet = 0 Then
        ICDAccess_ReadSectorsRAW = STREAM_END_OF_STREAM
    Else
        ICDAccess_ReadSectorsRAW = STREAM_OK
    End If
End Function

Private Function ICDAccess_SelectDevice( _
    ByVal index As Long _
) As STREAM_ERR

    If index < 0 Or index > lngDriveCount - 1 Then
        ICDAccess_SelectDevice = STREAM_OUT_OF_RANGE
        Exit Function
    End If

    lngSelected = index
    ICDAccess_SelectDevice = STREAM_OK
End Function
