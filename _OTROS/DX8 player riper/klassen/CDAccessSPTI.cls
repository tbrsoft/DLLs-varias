VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDAccessSPTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Low Level CD access through SCSI Pass Through
'
' + full access to the drive
' - only NT/2K/XP, needs admin privs

Implements ICDAccess

Private Type SPTIDrive
    char                As String
    handle              As hFile
    vendor              As String
    product             As String
    revision            As String
End Type

Private Type SPTIResult
    retval              As Byte
    sense               As SENSE_DATA_FMT
End Type

Private udtDrives()     As SPTIDrive

Private lngDriveCount   As Long
Private lngSelected     As Long

Private blnInitialized  As Boolean

Private Sub GetDeviceName( _
    ByVal index As Long _
)

    Dim udtInq  As SCSI_INQUIRY
    Dim udtRet  As SPTIResult
    Dim cdb(5)  As Byte

    cdb(0) = SCSI_INQUIRY
    cdb(4) = Len(udtInq)

    udtRet = SPTICmd(udtDrives(index).handle, _
                     cdb, 6, SCSI_IOCTL_DATA_IN, _
                     VarPtr(udtInq), Len(udtInq))

    If udtRet.retval <> 1 Then
        Debug.Print "SPTI: Inquiry fehlgeschlagen"
        Debug.Print "Sense Key: " & udtRet.sense.SenseKey
        Debug.Print "Asc: " & udtRet.sense.AddSenseCode
        Debug.Print "Ascq: " & udtRet.sense.AddSenQual
    End If

    With udtDrives(index)
        .vendor = RemNullChars(StrConv(udtInq.vendor, vbUnicode))
        .product = RemNullChars(StrConv(udtInq.product, vbUnicode))
        .revision = RemNullChars(StrConv(udtInq.revision, vbUnicode))
    End With
End Sub

Private Sub GetDriveList()
    Dim i       As Integer
    Dim hDrive  As hFile

    For i = 2 To 26
        If GetDriveType(Chr$(i + 64) & ":") = DTYPE_CDROM Then
            hDrive = GetDriveHandle(Chr$(i + 64))   ' READ + WRITE
            If hDrive.handle <> INVALID_HANDLE Then
                ReDim Preserve udtDrives(lngDriveCount) As SPTIDrive

                udtDrives(lngDriveCount).handle = hDrive
                udtDrives(lngDriveCount).char = Chr$(i + 64)
                GetDeviceName lngDriveCount

                lngDriveCount = lngDriveCount + 1
            End If
        End If
    Next
End Sub

Private Sub Class_Terminate()
    ICDAccess_DeInitMe
End Sub

Private Function ICDAccess_DeInitMe() As STREAM_ERR
    Dim i   As Long

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

Private Function SPTICmd( _
    handle As hFile, _
    cdb() As Byte, _
    ByVal cdblen As Byte, _
    direction As Byte, _
    ptrbuf As Long, _
    buflen As Long _
) As SPTIResult

    Dim udtSPT      As SCSI_PASS_THROUGH_DIRECT_W_BUFFER
    Dim dwRet       As Long
    Dim dwBytesRet  As Long
    Dim udtRet      As SPTIResult

    With udtSPT
        .Length = SPT_SIZE_OF
        .TimeOutValue = 5
        .SenseInfoLength = SENSE_LEN
        .SenseInfoOffset = SPT_SENSE_OFFSET
        .DataIn = direction
        .DataBuffer = ptrbuf
        .DataTransferLength = buflen

        DXCopyMemory .cdb(0), cdb(0), cdblen
        .CdbLength = cdblen
    End With

    dwRet = DeviceIoControl(handle.handle, _
                            IOCTL_SCSI_PASS_THROUGH_DIRECT, _
                            udtSPT, Len(udtSPT), _
                            udtSPT, Len(udtSPT), _
                            dwBytesRet, ByVal 0&)

    With udtRet
        .retval = dwRet
        .sense = udtSPT.SenseBuffer
    End With

    SPTICmd = udtRet
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

Private Function ICDAccess_GetTOC( _
    refTOC As CDTOC _
) As STREAM_ERR

    Dim clsTOC  As CDTOC
    Dim udtTOC  As CDROM_TOC
    Dim retval  As SPTIResult
    Dim i       As Long
    Dim cdb(9)  As Byte

    Set clsTOC = New CDTOC

    If lngSelected = -1 Then
        ICDAccess_GetTOC = STREAM_INVALID_DEVICE
        Exit Function
    End If

    cdb(0) = SCSI_READ_TOC
    cdb(1) = &H0                ' 2 = MSF, 0 = LBA
    cdb(7) = HiByte(Len(udtTOC))
    cdb(8) = LoByte(Len(udtTOC))

    retval = SPTICmd(udtDrives(lngSelected).handle, cdb, 10, SCSI_IOCTL_DATA_IN, VarPtr(udtTOC), Len(udtTOC))
    If retval.retval <> 1 Then
        ICDAccess_GetTOC = STREAM_NOT_SUPPORTED
    Else
        If udtTOC.LastTrack = 0 Then
            ICDAccess_GetTOC = STREAM_NOT_SUPPORTED
            Exit Function
        End If

        clsTOC.TrackCount = udtTOC.LastTrack + 1

        For i = 1 To udtTOC.LastTrack + 1
            clsTOC.SetTrackData udtTOC.TrackData(i - 1)
        Next

        Set refTOC = clsTOC

        ICDAccess_GetTOC = STREAM_OK
    End If
End Function

Private Function ICDAccess_IsDiscPresent() As Boolean
    Dim cdb(5)  As Byte
    Dim retval  As SPTIResult

    If lngSelected = -1 Then
        Exit Function
    End If

    retval = SPTICmd(udtDrives(lngSelected).handle, cdb, 6, SCSI_IOCTL_DATA_IN, 0, 0)
    If retval.retval = 1 Then
        ICDAccess_IsDiscPresent = True
    Else
        If retval.sense.AddSenseCode = &H3A Then
            If retval.sense.AddSenQual = &H0 Then
                Debug.Print "ASPI: keine CD im LW"

            ' seems like ascq isn't returned by all drives
            ElseIf retval.sense.AddSenQual = &H1 Then
                Debug.Print "ASPI: keine CD im LW - LW geschlossen"
            ElseIf retval.sense.AddSenQual = &H2 Then
                Debug.Print "ASPI: keine CD im LW - LW geöffnet"
            End If
        End If
    End If
End Function

Private Function ICDAccess_CloseCD() As STREAM_ERR
    Dim cdb(5)  As Byte
    Dim retval  As SPTIResult

    If lngSelected = -1 Then
        ICDAccess_CloseCD = STREAM_INVALID_DEVICE
        Exit Function
    End If

    cdb(0) = SCSI_START_STP
    cdb(4) = &H3

    retval = SPTICmd(udtDrives(lngSelected).handle, cdb, 6, SCSI_IOCTL_DATA_IN, 0, 0)
    If retval.retval = 1 Then
        ICDAccess_CloseCD = STREAM_OK
    Else
        ICDAccess_CloseCD = STREAM_NOT_SUPPORTED
    End If
End Function

Private Function ICDAccess_OpenCD() As STREAM_ERR
    Dim cdb(5)  As Byte
    Dim retval  As SPTIResult

    If lngSelected = -1 Then
        ICDAccess_OpenCD = STREAM_INVALID_DEVICE
        Exit Function
    End If

    cdb(0) = SCSI_START_STP
    cdb(4) = &H2

    retval = SPTICmd(udtDrives(lngSelected).handle, cdb, 6, SCSI_IOCTL_DATA_IN, 0, 0)
    If retval.retval = 1 Then
        ICDAccess_OpenCD = STREAM_OK
    Else
        ICDAccess_OpenCD = STREAM_NOT_SUPPORTED
    End If
End Function

Private Function ICDAccess_ReadSectorsRAW( _
    ByVal LBA As Long, _
    ByVal sectors As Long, _
    ByVal ptr As Long _
) As STREAM_ERR

    Dim cdb(11) As Byte
    Dim ret     As SPTIResult

    cdb(0) = SCSI_READ_CD
    cdb(1) = SHR(1, 2)              ' expected: CDDA
    cdb(2) = SHR(LBA, 24) And &HFF
    cdb(3) = SHR(LBA, 16) And &HFF
    cdb(4) = SHR(LBA, 8) And &HFF
    cdb(5) = LBA And &HFF
    cdb(6) = SHR(sectors, 16) And &HFF
    cdb(7) = SHR(sectors, 8) And &HFF
    cdb(8) = sectors And &HFF
    cdb(9) = &HF8                   ' RAW

    ret = SPTICmd(udtDrives(lngSelected).handle, cdb, 12, SCSI_IOCTL_DATA_IN, ptr, sectors * 2352)
    If ret.retval <> 1 Then
        ' an error is simply the end of a track
        ICDAccess_ReadSectorsRAW = STREAM_END_OF_STREAM
    Else
        ICDAccess_ReadSectorsRAW = STREAM_OK
    End If
End Function

Private Sub Class_Initialize()
    lngSelected = -1
End Sub
