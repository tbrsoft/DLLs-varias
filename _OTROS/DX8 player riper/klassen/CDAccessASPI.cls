VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDAccessASPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Low Level CD access through ASPI
' (Advanced SCSI Programming Interface)
'
' + full access to drives
' - problems under NT/2K/XP

Implements ICDAccess

Private Const SRB_DIR_SCSI              As Long = &H0
Private Const SRB_POSTING               As Long = &H1
Private Const SRB_ENABLE_RESIDUAL_COUNT As Long = &H4
Private Const SRB_DIR_IN                As Long = &H8
Private Const SRB_DIR_OUT               As Long = &H10
Private Const SRB_EVENT_NOTIFY          As Long = &H40
Private Const RESIDUAL_COUNT_SUPPORTED  As Long = &H2
Private Const MAX_SRB_TIMEOUT           As Long = 1080001
Private Const DEFAULT_SRB_TIMEOUT       As Long = 1080001

Private Const SC_HA_INQUIRY             As Long = &H0
Private Const SC_GET_DEV_TYPE           As Long = &H1
Private Const SC_EXEC_SCSI_CMD          As Long = &H2
Private Const SC_ABORT_SRB              As Long = &H3
Private Const SC_RESET_DEV              As Long = &H4
Private Const SC_SET_HA_PARAMS          As Long = &H5
Private Const SC_GET_DISK_INFO          As Long = &H6
Private Const SC_RESCAN_SCSI_BUS        As Long = &H7
Private Const SC_GETSET_TIMEOUTS        As Long = &H8

Private Const SS_PENDING                As Long = &H0
Private Const SS_COMP                   As Long = &H1
Private Const SS_ABORTED                As Long = &H2
Private Const SS_ABORT_FAIL             As Long = &H3
Private Const SS_ERR                    As Long = &H4
Private Const SS_INVALID_CMD            As Long = &H80
Private Const SS_INVALID_HA             As Long = &H81
Private Const SS_NO_DEVICE              As Long = &H82
Private Const SS_INVALID_SRB            As Long = &HE0
Private Const SS_BUFFER_ALIGN           As Long = &HE1
Private Const SS_ILLEGAL_MODE           As Long = &HE2
Private Const SS_NO_ASPI                As Long = &HE3
Private Const SS_FAILED_INIT            As Long = &HE4
Private Const SS_ASPI_IS_BUSY           As Long = &HE5
Private Const SS_BUFFER_TOO_BIG         As Long = &HE6
Private Const SS_MISMATCHED_COMPONENTS  As Long = &HE7
Private Const SS_NO_ADAPTERS            As Long = &HE8
Private Const SS_INSUFFICIENT_RESOURCES As Long = &HE9
Private Const SS_ASPI_IS_SHUTDOWN       As Long = &HEA
Private Const SS_BAD_INSTALL            As Long = &HEB

Private Const HASTAT_OK                 As Long = &H0
Private Const HASTAT_SEL_TO             As Long = &H11
Private Const HASTAT_DO_DU              As Long = &H12
Private Const HASTAT_BUS_FREE           As Long = &H13
Private Const HASTAT_PHASE_ERR          As Long = &H14
Private Const HASTAT_TIMEOUT            As Long = &H9
Private Const HASTAT_COMMAND_TIMEOUT    As Long = &HB
Private Const HASTAT_MESSAGE_REJECT     As Long = &HD
Private Const HASTAT_BUS_RESET          As Long = &HE
Private Const HASTAT_PARITY_ERROR       As Long = &HF
Private Const HASTAT_REQUEST_SENSE_FAILED As Long = &H10

Private Type SRB_HAINQUIRY
    SRB_Cmd                             As Byte
    SRB_Status                          As Byte
    SRB_HaId                            As Byte
    SRB_Flags                           As Byte
    SRB_Hdr_Rsvd                        As Long
    HA_Count                            As Byte
    HA_SCSI_ID                          As Byte
    HA_ManagerId(15)                    As Byte
    HA_Identifier(15)                   As Byte
    HA_Unique(15)                       As Byte
    HA_Rsvd1                            As Integer
    pad(19)                             As Byte
End Type

Private Type SRB_GetDiskInfo
    SRB_Cmd                             As Byte
    SRB_Status                          As Byte
    SRB_HaId                            As Byte
    SRB_Flags                           As Byte
    SRB_Hdr_Rsvd                        As Long
    SRB_Target                          As Byte
    SRB_LUN                             As Byte
    SRB_DriveFlags                      As Byte
    SRB_Int13DrvInfo                    As Byte
    SRB_Heads                           As Byte
    SRB_Sectors                         As Byte
    SRB_Rsvd1(9)                        As Byte
End Type

Private Type SRB_GetDevType
    SRB_Cmd                             As Byte
    SRB_Status                          As Byte
    SRB_HaId                            As Byte
    SRB_Flags                           As Byte
    SRB_Hdr_Rsvd                        As Long
    SRB_Target                          As Byte
    SRB_LUN                             As Byte
    SRB_DeviceType                      As Byte
    SRB_Rsvd1                           As Byte
    pad(67)                             As Byte
End Type

Private Type SRB_ExecSCSICmd
    SRB_Cmd                             As Byte
    SRB_Status                          As Byte
    SRB_HaId                            As Byte
    SRB_Flags                           As Byte
    SRB_Hdr_Rsvd                        As Long
    SRB_Target                          As Byte
    SRB_LUN                             As Byte
    SRB_Rsvd1                           As Integer
    SRB_BufLen                          As Long
    SRB_BufPointer                      As Long
    SRB_SenseLen                        As Byte
    SRB_CDBLen                          As Byte
    SRB_HaStat                          As Byte
    SRB_TargStat                        As Byte
    SRB_PostProc                        As Long
    SRB_Rsvd2(19)                       As Byte
    CDBByte(15)                         As Byte
    SenseArea(SENSE_LEN + 2)            As Byte
End Type

Private Type SRB_BusDeviceReset
    SRB_Cmd                             As Byte
    SRB_Status                          As Byte
    SRB_HaId                            As Byte
    SRB_Flags                           As Byte
    SRB_Hdr_Rsvd                        As Long
    SRB_Target                          As Byte
    SRB_LUN                             As Byte
    SRB_Rsvd1(11)                       As Byte
    SRB_HaStat                          As Byte
    SRB_TargStat                        As Byte
    SRB_PostProc                        As Long
    SRB_Rsvd2(35)                       As Byte
    pad(15)                             As Byte
End Type

Private Type ASPI32BUFF
    AB_BufPointer                       As Long
    AB_BufLen                           As Long
    AB_ZeroFill                         As Long
    AB_Reserved                         As Long
End Type

Private Type ASPIDrive
    HA                                  As Byte
    ID                                  As Byte
    LUN                                 As Byte
    vendor                              As String
    product                             As String
    revision                            As String
End Type

Private Type ASPIResult
    retval                              As Byte
    sense                               As SENSE_DATA_FMT
End Type

Private clsDLL                          As clsDLL
Private Const FNC_INFO                  As String = "GetASPI32SupportInfo"
Private Const FNC_CMD                   As String = "SendASPI32Command"

Private udtDrives()                     As ASPIDrive
Private lngDriveCount                   As Long

Private lngSelected                     As Long

Private blnInitialized                  As Boolean

Private Sub Class_Initialize()
    lngSelected = -1
End Sub

Private Function ICDAccess_DeviceChar( _
    ByVal index As Long _
) As String

    Dim i           As Integer
    Dim udtDevInfo  As SRB_GetDiskInfo

    ' NT/2k/XP
    If GetWindowsVersion.dwPlatformId = VER_PLATFORM_WIN32_NT Then

        For i = 2 To 26

            With GetBusAddr(Chr$(i + 64))
                If .PortNumber = udtDrives(index).HA Then
                    If .TargetId = udtDrives(index).ID Then
                        If .LUN = udtDrives(index).LUN Then
                            ICDAccess_DeviceChar = Chr$(i + 64)
                            Exit Function
                        End If
                    End If
                End If
            End With

        Next

    ' 95/98/Me
    Else

        udtDevInfo.SRB_Cmd = SC_GET_DISK_INFO
        udtDevInfo.SRB_HaId = udtDrives(lngSelected).HA
        udtDevInfo.SRB_Target = udtDrives(lngSelected).ID
        udtDevInfo.SRB_LUN = udtDrives(lngSelected).LUN

        clsDLL.CallFunc FNC_CMD, VarPtr(udtDevInfo)

        ICDAccess_DeviceChar = Chr$(udtDevInfo.SRB_Int13DrvInfo + 65)

    End If
End Function

Private Function ICDAccess_GetTOC( _
    refTOC As CDTOC _
) As STREAM_ERR

    Dim clsTOC  As CDTOC
    Dim udtTOC  As CDROM_TOC
    Dim retval  As ASPIResult
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

    retval = ASPICmd(udtDrives(lngSelected), cdb, 10, SRB_DIR_IN, VarPtr(udtTOC), Len(udtTOC))
    If retval.retval <> SS_COMP Then
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
    Dim retval  As ASPIResult

    If lngSelected = -1 Then
        Exit Function
    End If

    retval = ASPICmd(udtDrives(lngSelected), cdb, 6, SRB_DIR_IN, 0, 0)
    If retval.retval = SS_COMP Then
        ICDAccess_IsDiscPresent = True
    Else
        If retval.sense.AddSenseCode = &H3A Then
            If retval.sense.AddSenQual = &H0 Then
                Debug.Print "ASPI: keine CD im LW"

            ' ascq isn't returned by all drives
            ' (seems like)
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
    Dim retval  As ASPIResult

    If lngSelected = -1 Then
        ICDAccess_CloseCD = STREAM_INVALID_DEVICE
        Exit Function
    End If

    cdb(0) = SCSI_START_STP
    cdb(4) = &H3

    retval = ASPICmd(udtDrives(lngSelected), cdb, 6, SRB_DIR_IN, 0, 0)
    If retval.retval = SS_COMP Then
        ICDAccess_CloseCD = STREAM_OK
    Else
        ICDAccess_CloseCD = STREAM_NOT_SUPPORTED
    End If
End Function

Private Function ICDAccess_DeInitMe() As STREAM_ERR
    blnInitialized = False
    ICDAccess_DeInitMe = STREAM_OK
End Function

Private Sub GetDeviceName( _
    ByVal index As Long _
)

    Dim udtInq  As SCSI_INQUIRY
    Dim ret     As ASPIResult
    Dim cdb(5)  As Byte

    If Not blnInitialized Then Exit Sub

    cdb(0) = SCSI_INQUIRY
    cdb(4) = Len(udtInq)

    ret = ASPICmd(udtDrives(index), cdb, 6, SRB_DIR_IN, VarPtr(udtInq), Len(udtInq))
    If ret.retval <> SS_COMP Then
        Debug.Print "ASPI - GetDeviceName - Error"
        Debug.Print "Sense Key:  " & Hex$(ret.sense.SenseKey)
        Debug.Print "Add Sense:  " & Hex$(ret.sense.AddSenLen)
        Debug.Print "Add SenseQ: " & Hex$(ret.sense.AddSenQual)
        Exit Sub
    End If

    With udtDrives(index)
        .vendor = RemNullChars(StrConv(udtInq.vendor, vbUnicode))
        .product = RemNullChars(StrConv(udtInq.product, vbUnicode))
        .revision = RemNullChars(StrConv(udtInq.revision, vbUnicode))
    End With
End Sub

Private Sub GetDriveList()
    Dim udtHAInq    As SRB_HAINQUIRY
    Dim udtDevType  As SRB_GetDevType
    Dim maxtgt      As Long
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long

    lngDriveCount = 0
    If Not blnInitialized Then Exit Sub

    For i = 0 To GetNumAdapters() - 1
        udtHAInq.SRB_Cmd = SC_HA_INQUIRY
        udtHAInq.SRB_HaId = i
        clsDLL.CallFunc FNC_CMD, VarPtr(udtHAInq)

        If udtHAInq.SRB_Status = SS_COMP Then

            maxtgt = udtHAInq.HA_Unique(3)
            If maxtgt = 0 Then maxtgt = MAXTARG

            For j = 0 To maxtgt - 1

                For k = 0 To MAXLUN - 1

                    With udtDevType
                        .SRB_Cmd = SC_GET_DEV_TYPE
                        .SRB_HaId = i
                        .SRB_Target = j
                        .SRB_LUN = k
                    End With

                    clsDLL.CallFunc FNC_CMD, VarPtr(udtDevType)

                    If udtDevType.SRB_Status = SS_COMP Then

                        If udtDevType.SRB_DeviceType = DTYPE_CDROM Then

                            ReDim Preserve udtDrives(lngDriveCount) As ASPIDrive

                            With udtDrives(lngDriveCount)
                                .HA = i
                                .ID = j
                                .LUN = k
                            End With

                            GetDeviceName lngDriveCount

                            lngDriveCount = lngDriveCount + 1

                        End If

                    End If

                Next

            Next

        End If
    Next
End Sub

Private Function ICDAccess_DeviceCount() As Long
    ICDAccess_DeviceCount = lngDriveCount
End Function

Private Function GetNumAdapters() As Long
    GetNumAdapters = LoByte(LoWord(clsDLL.CallFunc(FNC_INFO)))
End Function

Private Function ICDAccess_DeviceName( _
    ByVal index As Long _
) As String

    With udtDrives(index)
        ICDAccess_DeviceName = .vendor & " " & .product & " " & .revision
    End With
End Function

Private Function ICDAccess_InitMe() As STREAM_ERR
    Dim lngRet As Long

    ICDAccess_InitMe = STREAM_NOT_SUPPORTED

    Set clsDLL = New clsDLL

    If Not clsDLL.LoadDLL("wnaspi32.dll") Then
        Exit Function
    End If

    clsDLL.IsCDECL = True

    lngRet = clsDLL.CallFunc(FNC_INFO)

    If HiByte(LoWord(lngRet)) = SS_COMP Then
        blnInitialized = True
        GetDriveList
        If lngDriveCount > 0 Then
            ICDAccess_InitMe = STREAM_OK
        End If
    End If
End Function

Private Function ASPICmd( _
    drive As ASPIDrive, _
    cdb() As Byte, _
    cdblen As Byte, _
    Flags As Byte, _
    ptrbuf As Long, _
    buflen As Long _
) As ASPIResult

    Dim udtExec As SRB_ExecSCSICmd
    Dim i       As Long

    With udtExec
        .SRB_Cmd = SC_EXEC_SCSI_CMD
        .SRB_HaId = drive.HA
        .SRB_Target = drive.ID
        .SRB_LUN = drive.LUN
        .SRB_BufLen = buflen
        .SRB_BufPointer = ptrbuf
        .SRB_SenseLen = SENSE_LEN
        .SRB_Flags = Flags
        .SRB_CDBLen = cdblen
        DXCopyMemory .CDBByte(0), cdb(0), cdblen
    End With

    clsDLL.CallFunc FNC_CMD, VarPtr(udtExec)

    Do While udtExec.SRB_Status = SS_PENDING
        'DoEvents
    Loop

    ASPICmd.retval = udtExec.SRB_Status
    DXCopyMemory ASPICmd.sense, udtExec.SenseArea(0), SENSE_LEN
End Function

Private Function ICDAccess_OpenCD() As STREAM_ERR
    Dim cdb(5)  As Byte
    Dim retval  As ASPIResult

    If lngSelected = -1 Then
        ICDAccess_OpenCD = STREAM_INVALID_DEVICE
        Exit Function
    End If

    cdb(0) = SCSI_START_STP
    cdb(4) = &H2

    retval = ASPICmd(udtDrives(lngSelected), cdb, 6, SRB_DIR_IN, 0, 0)
    If retval.retval = SS_COMP Then
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
    Dim ret     As ASPIResult

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

    ret = ASPICmd(udtDrives(lngSelected), cdb, 12, SRB_DIR_IN, ptr, sectors * 2352)
    If ret.retval <> SS_COMP Then
        ' an error is simply the end of a track
        ICDAccess_ReadSectorsRAW = STREAM_END_OF_STREAM
    Else
        ICDAccess_ReadSectorsRAW = STREAM_OK
    End If
End Function

Private Function ICDAccess_SelectDevice( _
    ByVal index As Long _
) As STREAM_ERR

    If index > lngDriveCount - 1 Or index < 0 Then
        ICDAccess_SelectDevice = STREAM_OUT_OF_RANGE
        Exit Function
    End If

    lngSelected = index
    ICDAccess_SelectDevice = STREAM_OK
End Function
