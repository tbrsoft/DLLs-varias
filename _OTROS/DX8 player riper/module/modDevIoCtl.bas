Attribute VB_Name = "modDevIoCtl"
Option Explicit

' DeviceIoControl constants/structures
'
' translated a lot from the Wine project

Public Declare Function DeviceIoControl Lib "kernel32" ( _
    ByVal hDevice As Long, _
    ByVal dwIoControlCode As Long, _
    lpInBuffer As Any, _
    ByVal nInBufferSize As Long, _
    lpOutBuffer As Any, _
    ByVal nOutBufferSize As Long, _
    lpBytesReturned As Long, _
    lpOverlapped As Any _
) As Long

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Const FILE_ANY_ACCESS                As Long = 0
Public Const FILE_READ_ACCESS               As Long = &H1
Public Const FILE_WRITE_ACCESS              As Long = &H2

Public Const METHOD_BUFFERED                As Long = 0
Public Const METHOD_IN_DIRECT               As Long = 1
Public Const METHOD_OUT_DIRECT              As Long = 2
Public Const METHOD_NEITHER                 As Long = 3

Public Const SCSI_IOCTL_DATA_OUT            As Long = 0
Public Const SCSI_IOCTL_DATA_IN             As Long = 1
Public Const SCSI_IOCTL_DATA_UNSPECIFIED    As Long = 2

Public Const FILE_DEVICE_BEEP               As Long = &H1
Public Const FILE_DEVICE_CD_ROM             As Long = &H2
Public Const FILE_DEVICE_CD_ROM_FILE_SYSTEM As Long = &H3
Public Const FILE_DEVICE_CONTROLLER         As Long = &H4
Public Const FILE_DEVICE_DATALINK           As Long = &H5
Public Const FILE_DEVICE_DFS                As Long = &H6
Public Const FILE_DEVICE_DISK               As Long = &H7
Public Const FILE_DEVICE_DISK_FILE_SYSTEM   As Long = &H8
Public Const FILE_DEVICE_FILE_SYSTEM        As Long = &H9
Public Const FILE_DEVICE_INPORT_PORT        As Long = &HA
Public Const FILE_DEVICE_KEYBOARD           As Long = &HB
Public Const FILE_DEVICE_MAILSLOT           As Long = &HC
Public Const FILE_DEVICE_MIDI_IN            As Long = &HD
Public Const FILE_DEVICE_MIDI_OUT           As Long = &HE
Public Const FILE_DEVICE_MOUSE              As Long = &HF
Public Const FILE_DEVICE_MULTI_UNC_PROVIDER As Long = &H10
Public Const FILE_DEVICE_NAMED_PIPE         As Long = &H11
Public Const FILE_DEVICE_NETWORK            As Long = &H12
Public Const FILE_DEVICE_NETWORK_BROWSER    As Long = &H13
Public Const FILE_DEVICE_NETWORK_FILE_SYSTEM As Long = &H14
Public Const FILE_DEVICE_NULL               As Long = &H15
Public Const FILE_DEVICE_PARALLEL_PORT      As Long = &H16
Public Const FILE_DEVICE_PHYSICAL_NETCARD   As Long = &H17
Public Const FILE_DEVICE_PRINTER            As Long = &H18
Public Const FILE_DEVICE_SCANNER            As Long = &H19
Public Const FILE_DEVICE_SERIAL_MOUSE_PORT  As Long = &H1A
Public Const FILE_DEVICE_SERIAL_PORT        As Long = &H1B
Public Const FILE_DEVICE_SCREEN             As Long = &H1C
Public Const FILE_DEVICE_SOUND              As Long = &H1D
Public Const FILE_DEVICE_STREAMS            As Long = &H1E
Public Const FILE_DEVICE_TAPE               As Long = &H1F
Public Const FILE_DEVICE_TAPE_FILE_SYSTEM   As Long = &H20
Public Const FILE_DEVICE_TRANSPORT          As Long = &H21
Public Const FILE_DEVICE_UNKNOWN            As Long = &H22
Public Const FILE_DEVICE_VIDEO              As Long = &H23
Public Const FILE_DEVICE_VIRTUAL_DISK       As Long = &H24
Public Const FILE_DEVICE_WAVE_IN            As Long = &H25
Public Const FILE_DEVICE_WAVE_OUT           As Long = &H26
Public Const FILE_DEVICE_8042_PORT          As Long = &H27
Public Const FILE_DEVICE_NETWORK_REDIRECTOR As Long = &H28
Public Const FILE_DEVICE_BATTERY            As Long = &H29
Public Const FILE_DEVICE_BUS_EXTENDER       As Long = &H2A
Public Const FILE_DEVICE_MODEM              As Long = &H2B
Public Const FILE_DEVICE_VDM                As Long = &H2C
Public Const FILE_DEVICE_MASS_STORAGE       As Long = &H2D
Public Const FILE_DEVICE_SMB                As Long = &H2E
Public Const FILE_DEVICE_KS                 As Long = &H2F
Public Const FILE_DEVICE_CHANGER            As Long = &H30
Public Const FILE_DEVICE_SMARTCARD          As Long = &H31
Public Const FILE_DEVICE_ACPI               As Long = &H32
Public Const FILE_DEVICE_DVD                As Long = &H33
Public Const FILE_DEVICE_FULLSCREEN_VIDEO   As Long = &H34
Public Const FILE_DEVICE_DFS_FILE_SYSTEM    As Long = &H35
Public Const FILE_DEVICE_DFS_VOLUME         As Long = &H36
Public Const FILE_DEVICE_SERENUM            As Long = &H37
Public Const FILE_DEVICE_TERMSRV            As Long = &H38
Public Const FILE_DEVICE_KSEC               As Long = &H39

Public Const IOCTL_SCSI_BASE                As Long = FILE_DEVICE_CONTROLLER
Public Const IOCTL_STORAGE_BASE             As Long = FILE_DEVICE_MASS_STORAGE
Public Const IOCTL_DISK_BASE                As Long = FILE_DEVICE_DISK

Public Const MAXIMUM_NUMBER_TRACKS          As Long = 100
Public Const MAXIMUM_CDROM_SIZE             As Long = 804
Public Const MINIMUM_CDROM_READ_TOC_EX_SIZE As Long = 2

Public Const CDROM_DISK_AUDIO_TRACK         As Long = 1
Public Const CDROM_DISK_DATA_TRACK          As Long = 2

Public Const IOCTL_CDROM_SUB_Q_CHANNEL      As Long = 0
Public Const IOCTL_CDROM_CURRENT_POSITION   As Long = 1
Public Const IOCTL_CDROM_MEDIA_CATALOG      As Long = 2
Public Const IOCTL_CDROM_TRACK_ISRC         As Long = 3

Public Const ADR_NO_MODE_INFORMATION        As Long = 0
Public Const ADR_ENCODES_CURRENT_POSITION   As Long = 1
Public Const ADR_ENCODES_MEDIA_CATALOG      As Long = 2
Public Const ADR_ENCODES_ISRC               As Long = 3

Public Const AUDIO_STATUS_NOT_SUPPORTED     As Long = 0
Public Const AUDIO_STATUS_IN_PROGRESS       As Long = &H11
Public Const AUDIO_STATUS_PAUSED            As Long = &H12
Public Const AUDIO_STATUS_PLAY_COMPLETE     As Long = &H13
Public Const AUDIO_STATUS_PLAY_ERROR        As Long = &H14
Public Const AUDIO_STATUS_NO_STATUS         As Long = &H15

Public Const AUDIO_WITH_PREEMPHASIS         As Long = &H1
Public Const DIGITAL_COPY_PERMITTED         As Long = &H2
Public Const AUDIO_DATA_TRACK               As Long = &H4
Public Const TWO_FOUR_CHANNEL_AUDIO         As Long = &H8

Public IOCTL_SCSI_PASS_THROUGH              As Long
Public IOCTL_SCSI_MINIPORT                  As Long
Public IOCTL_SCSI_GET_INQUIRY_DATA          As Long
Public IOCTL_SCSI_GET_CAPABILITIES          As Long
Public IOCTL_SCSI_PASS_THROUGH_DIRECT       As Long
Public IOCTL_SCSI_GET_ADDRESS               As Long
Public IOCTL_SCSI_RESCAN_BUS                As Long
Public IOCTL_SCSI_GET_DUMP_POINTERS         As Long
Public IOCTL_SCSI_FREE_DUMP_POINTERS        As Long
Public IOCTL_IDE_PASS_THROUGH               As Long

Public IOCTL_CDROM_RAW_READ                 As Long
Public IOCTL_CDROM_READ_TOC                 As Long
Public IOCTL_CDROM_GET_CONTROL              As Long
Public IOCTL_CDROM_PLAY_AUDIO_MSF           As Long
Public IOCTL_CDROM_SEEK_AUDIO_MSF           As Long
Public IOCTL_CDROM_STOP_AUDIO               As Long
Public IOCTL_CDROM_PAUSE_AUDIO              As Long
Public IOCTL_CDROM_RESUME_AUDIO             As Long
Public IOCTL_CDROM_GET_VOLUME               As Long
Public IOCTL_CDROM_SET_VOLUME               As Long
Public IOCTL_CDROM_READ_Q_CHANNEL           As Long
Public IOCTL_CDROM_GET_LAST_SESSION         As Long
Public IOCTL_CDROM_DISK_TYPE                As Long

Public IOCTL_STORAGE_CHECK_VERIFY           As Long
Public IOCTL_STORAGE_MEDIA_REMOVAL          As Long
Public IOCTL_STORAGE_EJECT_MEDIA            As Long
Public IOCTL_STORAGE_LOAD_MEDIA             As Long
Public IOCTL_STORAGE_RESERVE                As Long
Public IOCTL_STORAGE_RELEASE                As Long
Public IOCTL_STORAGE_FIND_NEW_DEVICES       As Long
Public IOCTL_STORAGE_EJECTION_CONTROL       As Long
Public IOCTL_STORAGE_MCN_CONTROL            As Long

Public IOCTL_STORAGE_GET_MEDIA_TYPES        As Long
Public IOCTL_STORAGE_GET_MEDIA_TYPES_EX     As Long

Public IOCTL_STORAGE_RESET_BUS              As Long
Public IOCTL_STORAGE_RESET_DEVICE           As Long
Public IOCTL_STORAGE_GET_DEVICE_NUMBER      As Long

Public IOCTL_DISK_GET_DRIVE_GEOMETRY        As Long
Public IOCTL_DISK_GET_PARTITION_INFO        As Long
Public IOCTL_DISK_SET_PARTITION_INFO        As Long
Public IOCTL_DISK_GET_DRIVE_LAYOUT          As Long
Public IOCTL_DISK_SET_DRIVE_LAYOUT          As Long
Public IOCTL_DISK_VERIFY                    As Long
Public IOCTL_DISK_FORMAT_TRACKS             As Long
Public IOCTL_DISK_REASSIGN_BLOCKS           As Long
Public IOCTL_DISK_PERFORMANCE               As Long
Public IOCTL_DISK_IS_WRITABLE               As Long
Public IOCTL_DISK_LOGGING                   As Long
Public IOCTL_DISK_FORMAT_TRACKS_EX          As Long
Public IOCTL_DISK_HISTOGRAM_STRUCTURE       As Long
Public IOCTL_DISK_HISTOGRAM_DATA            As Long
Public IOCTL_DISK_HISTOGRAM_RESET           As Long
Public IOCTL_DISK_REQUEST_STRUCTURE         As Long
Public IOCTL_DISK_REQUEST_DATA              As Long
Public IOCTL_DISK_CHECK_VERIFY              As Long
Public IOCTL_DISK_MEDIA_REMOVAL             As Long
Public IOCTL_DISK_EJECT_MEDIA               As Long
Public IOCTL_DISK_LOAD_MEDIA                As Long
Public IOCTL_DISK_RESERVE                   As Long
Public IOCTL_DISK_RELEASE                   As Long
Public IOCTL_DISK_FIND_NEW_DEVICES          As Long
Public IOCTL_DISK_GET_MEDIA_TYPES           As Long

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Type TRACK_DATA
    Rsvd                    As Byte
    ADRCTL                  As Byte
    TrackNumber             As Byte
    rsvd1                   As Byte
    addr(3)                 As Byte
End Type

Public Type CDROM_DISK_DATA
    DiskData                As Long
End Type

Public Type CDROM_PLAY_AUDIO_MSF
    StartingM               As Byte
    StartingS               As Byte
    StartingF               As Byte
    EndingM                 As Byte
    EndingS                 As Byte
    EndingF                 As Byte
End Type

Public Type CDROM_READ_TOC_EX
    format                  As Byte ' fmt: 4, rsvd: 3, msf:1
    SessionTrack            As Byte
    rsvd2                   As Byte
    rsvd3                   As Byte
End Type

Public Type CDROM_SEEK_AUDIO_MSF
    m                       As Byte
    s                       As Byte
    F                       As Byte
End Type

Public Type CDROM_SUB_Q_DATA_FORMAT
    format                  As Byte
    track                   As Byte
End Type

Public Type CDROM_TOC
    Length(1)               As Byte
    FirstTrack              As Byte
    LastTrack               As Byte
    TrackData(MAXIMUM_NUMBER_TRACKS - 1) As TRACK_DATA
End Type

Public Type CDROM_TOC_ATIP_DATA_BLOCK
    WritePower              As Byte ' RefSpeed: 3, Rsvd3: 1, WritePow: 3, True1: 1
    Rsvd                    As Byte ' rsvd4: 6, UnrestrictedUse: 1, rsvd5: 1
    IsCDRW                  As Byte ' A3Valid: 1, A2Valid: 1, A1Valid: 1, rsvd6: 3, IsCDRW: 1, True2: 1
    rsvd7                   As Byte
    LeadInMSF(2)            As Byte
    rsvd8                   As Byte
    LeadOutMSF(2)           As Byte
    rsvd9                   As Byte
    A1Values(2)             As Byte
    rsvd10                  As Byte
    A2Values(2)             As Byte
    rsvd11                  As Byte
    A3Values(2)             As Byte
    rsvd12                  As Byte
End Type

Public Type CDROM_TOC_ATIP_DATA
    Length(1)               As Byte
    Rsvd(1)                 As Byte
    Descriptor              As CDROM_TOC_ATIP_DATA_BLOCK
End Type

Public Type CDROM_TOC_CD_TEXT_DATA_BLOCK
    PackType                As Byte
    TrackNumber             As Byte ' TrackNum: 7, ExtensionFlag: 1
    SequenceNumber          As Byte
    CharacterPosition       As Byte ' CharPos: 4, BlockNumber: 3, Unicode: 1
    Text(11)                As Byte
    crc(1)                  As Byte
End Type

Public Type CDROM_TOC_CD_TEXT_DATA
    Length(1)               As Byte
    Rsvd(1)                 As Byte
    Descriptor(255)         As CDROM_TOC_CD_TEXT_DATA_BLOCK
End Type

Public Type CDROM_TOC_FULL_TOC_DATA_BLOCK
    SessionNumber           As Byte
    ADRCTL                  As Byte
    rsvd1                   As Byte
    Point                   As Byte
    MsfExtra(2)             As Byte
    Zero                    As Byte
    MSF(2)                  As Byte
End Type

Public Type CDROM_TOC_FULL_TOC_DATA
    Length(1)               As Byte
    FirstCompleteSess       As Byte
    LastCompleteSess        As Byte
    Descriptor(255)         As CDROM_TOC_FULL_TOC_DATA_BLOCK
End Type

Public Type CDROM_TOC_PMA_DATA
    Length(1)               As Byte
    Rsvd(1)                 As Byte
    Descriptor(255)         As CDROM_TOC_FULL_TOC_DATA_BLOCK
End Type

Public Type SUB_Q_HEADER
    Rsvd                    As Byte
    AudioStatus             As Byte
    datalength(1)           As Byte
End Type

Public Type SUB_Q_MEDIA_CATALOG_NUMBER
    header                  As SUB_Q_HEADER
    FormatCode              As Byte
    Rsvd(2)                 As Byte
    Mcval                   As Byte ' Rsvd1: 7, Mcval: 1
    MediaCatalog(14)        As Byte
End Type

Public Type SUB_Q_TRACK_ISRC
    header                  As SUB_Q_HEADER
    FormatCode              As Byte
    rsvd0                   As Byte
    track                   As Byte
    rsvd1                   As Byte
    Tcval                   As Byte ' Rsvd2: 7, Tcval: 1
    TrackIsrc(14)           As Byte
End Type

Public Type SUB_Q_CURRENT_POSITION
    header                  As SUB_Q_HEADER
    FormatCode              As Byte
    ADRCTL                  As Byte
    TrackNumber             As Byte
    IndexNumber             As Byte
    AbsoluteAddr(3)         As Byte
    TrackRelAddr(3)         As Byte
End Type

Public Type SUB_Q_CHANNEL_DATA
    CurrentPosition         As SUB_Q_CURRENT_POSITION
    MediaCatalog            As SUB_Q_MEDIA_CATALOG_NUMBER
    TrackIsrc               As SUB_Q_TRACK_ISRC
End Type

Public Type CDROM_AUDIO_CONTROL
    LbaFormat               As Byte
    LogicalBlocksPerSecond  As Integer
End Type

Public Type VOLUME_CONTROL
    PortVolume(3)           As Byte
End Type

Public Type DISK_GEOMETRY
    Cylinders               As LARGE_INTEGER
    MediaType               As Long
    TracksPerCylinder       As Long
    SectorsPerTrack         As Long
    BytesPerSector          As Long
End Type

Public Type RAW_READ_INFO
    DiskOffset              As LARGE_INTEGER
    SectorCount             As Long
    TrackMode               As TRACK_MODE_TYPE
End Type

Public Type INQUIRY_DATA
    inqDevType              As Byte
    inqRMB                  As Byte
    inqVersion              As Byte
    inqAtapiVersion         As Byte
    inqLength               As Byte
    inqReserved(2)          As Byte
    inqVendor(7)            As Byte
    inqProdID(15)           As Byte
    inqRev(3)               As Byte
    inqReserved2(59)        As Byte
End Type

Public Type SCSI_ADDRESS
    Length                  As Long
    PortNumber              As Byte
    PathId                  As Byte
    TargetId                As Byte
    LUN                     As Byte
End Type

Public Type SCSI_PASS_THROUGH_DIRECT_W_BUFFER
    Length                  As Integer
    ScsiStatus              As Byte
    PathId                  As Byte
    TargetId                As Byte
    LUN                     As Byte
    CdbLength               As Byte
    SenseInfoLength         As Byte
    DataIn                  As Byte
    DataTransferLength      As Long
    TimeOutValue            As Long
    DataBuffer              As Long
    SenseInfoOffset         As Long
    cdb(15)                 As Byte
    SenseBuffer             As SENSE_DATA_FMT
End Type

Public Const SPT_SIZE_OF        As Long = 44
Public Const SPT_SENSE_OFFSET   As Long = 44
Public Const SENSE_LEN          As Long = 14

Private Type SCSI_BUS_DATA
    NumberOfLogicalUnits    As Byte
    InitiatorBusID          As Byte
    InquiryDataOffset       As Long
End Type

Private Type SCSI_ADAPTER_BUS_INFO
    NumberOfBusses          As Byte
    BusData                 As SCSI_BUS_DATA
End Type

Private Type SCSI_INQUIRY_DATA
    PathId                  As Byte
    TargetId                As Byte
    LUN                     As Byte
    DeviceClaimed           As Long
    InquiryDataLength       As Long
    NextInquiryDataOffset   As Long
    InquiryData             As Byte
End Type

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Enum TRACK_MODE_TYPE
    TRACK_MODE_YellowMode2
    TRACK_MODE_XAForm2
    TRACK_MODE_CDDA
End Enum

Public Enum CD_ROM_CD_TEXT_PACK
    CDROM_CD_TEXT_PACK_ALBUM_NAME = &H80
    CDROM_CD_TEXT_PACK_PERFORMER = &H81
    CDROM_CD_TEXT_PACK_SONGWRITER = &H82
    CDROM_CD_TEXT_PACK_COMPOSER = &H83
    CDROM_CD_TEXT_PACK_ARRANGER = &H84
    CDROM_CD_TEXT_PACK_MESSAGES = &H85
    CDROM_CD_TEXT_PACK_DISC_ID = &H86
    CDROM_CD_TEXT_PACK_GENRE = &H87
    CDROM_CD_TEXT_PACK_TOC_INFO = &H88
    CDROM_CD_TEXT_PACK_TOC_INFO2 = &H89
    CDROM_CD_TEXT_PACK_UPC_EAN = &H8E
    CDROM_CD_TEXT_PACK_SIZE_INFO = &H8F
End Enum

Public Enum READ_TOC_EX_FORMAT
    CDROM_READ_TOC_EX_FORMAT_TOC = 0
    CDROM_READ_TOC_EX_FORMAT_SESSION = 1
    CDROM_READ_TOC_EX_FORMAT_FULL_TOC = 2
    CDROM_READ_TOC_EX_FORMAT_PMA = 3
    CDROM_READ_TOC_EX_FORMAT_ATIP = 4
    CDROM_READ_TOC_EX_FORMAT_CDTEXT = 5
End Enum

Public Enum DEVICE_TYPES
    DTYPE_DASD = &H0          ' Disk Device
    DTYPE_SEQD = &H1          ' Tape Device
    DTYPE_PRNT = &H2          ' Printer
    DTYPE_PROC = &H3          ' Processor
    DTYPE_WORM = &H4          ' Write-once read-multiple
    DTYPE_CDROM = &H5         ' CD-ROM device
    DTYPE_SCAN = &H6          ' Scanner device
    DTYPE_OPTI = &H7          ' Optical memory device
    DTYPE_JUKE = &H8          ' Medium Changer device
    DTYPE_COMM = &H9          ' Communications device
    DTYPE_RESL = &HA          ' Reserved (low)
    DTYPE_RESH = &H1E         ' Reserved (high)
    DTYPE_UNKNOWN = &H1F      ' Unknown or no device type
End Enum

Public Sub InitIOCTLs()
          IOCTL_SCSI_PASS_THROUGH = CTL_CODE(IOCTL_SCSI_BASE, &H401, METHOD_BUFFERED, FILE_READ_ACCESS Or FILE_WRITE_ACCESS)
              IOCTL_SCSI_MINIPORT = CTL_CODE(IOCTL_SCSI_BASE, &H402, METHOD_BUFFERED, FILE_READ_ACCESS Or FILE_WRITE_ACCESS)
      IOCTL_SCSI_GET_INQUIRY_DATA = CTL_CODE(IOCTL_SCSI_BASE, &H403, METHOD_BUFFERED, FILE_ANY_ACCESS)
      IOCTL_SCSI_GET_CAPABILITIES = CTL_CODE(IOCTL_SCSI_BASE, &H404, METHOD_BUFFERED, FILE_ANY_ACCESS)
   IOCTL_SCSI_PASS_THROUGH_DIRECT = CTL_CODE(IOCTL_SCSI_BASE, &H405, METHOD_BUFFERED, FILE_READ_ACCESS Or FILE_WRITE_ACCESS)
           IOCTL_SCSI_GET_ADDRESS = CTL_CODE(IOCTL_SCSI_BASE, &H406, METHOD_BUFFERED, FILE_ANY_ACCESS)
            IOCTL_SCSI_RESCAN_BUS = CTL_CODE(IOCTL_SCSI_BASE, &H407, METHOD_BUFFERED, FILE_ANY_ACCESS)
     IOCTL_SCSI_GET_DUMP_POINTERS = CTL_CODE(IOCTL_SCSI_BASE, &H408, METHOD_BUFFERED, FILE_ANY_ACCESS)
    IOCTL_SCSI_FREE_DUMP_POINTERS = CTL_CODE(IOCTL_SCSI_BASE, &H409, METHOD_BUFFERED, FILE_ANY_ACCESS)
           IOCTL_IDE_PASS_THROUGH = CTL_CODE(IOCTL_SCSI_BASE, &H40A, METHOD_BUFFERED, FILE_READ_ACCESS Or FILE_WRITE_ACCESS)

             IOCTL_CDROM_RAW_READ = CTL_CODE(FILE_DEVICE_CD_ROM, &HF, METHOD_OUT_DIRECT, FILE_READ_ACCESS)
             IOCTL_CDROM_READ_TOC = CTL_CODE(FILE_DEVICE_CD_ROM, &H0, METHOD_BUFFERED, FILE_READ_ACCESS)
          IOCTL_CDROM_GET_CONTROL = CTL_CODE(FILE_DEVICE_CD_ROM, &HD, METHOD_BUFFERED, FILE_READ_ACCESS)
       IOCTL_CDROM_PLAY_AUDIO_MSF = CTL_CODE(FILE_DEVICE_CD_ROM, &H6, METHOD_BUFFERED, FILE_READ_ACCESS)
       IOCTL_CDROM_SEEK_AUDIO_MSF = CTL_CODE(FILE_DEVICE_CD_ROM, &H1, METHOD_BUFFERED, FILE_READ_ACCESS)
           IOCTL_CDROM_STOP_AUDIO = CTL_CODE(FILE_DEVICE_CD_ROM, &H2, METHOD_BUFFERED, FILE_READ_ACCESS)
          IOCTL_CDROM_PAUSE_AUDIO = CTL_CODE(FILE_DEVICE_CD_ROM, &H3, METHOD_BUFFERED, FILE_READ_ACCESS)
         IOCTL_CDROM_RESUME_AUDIO = CTL_CODE(FILE_DEVICE_CD_ROM, &H4, METHOD_BUFFERED, FILE_READ_ACCESS)
           IOCTL_CDROM_GET_VOLUME = CTL_CODE(FILE_DEVICE_CD_ROM, &H5, METHOD_BUFFERED, FILE_READ_ACCESS)
           IOCTL_CDROM_SET_VOLUME = CTL_CODE(FILE_DEVICE_CD_ROM, &HA, METHOD_BUFFERED, FILE_READ_ACCESS)
       IOCTL_CDROM_READ_Q_CHANNEL = CTL_CODE(FILE_DEVICE_CD_ROM, &HB, METHOD_BUFFERED, FILE_READ_ACCESS)
     IOCTL_CDROM_GET_LAST_SESSION = CTL_CODE(FILE_DEVICE_CD_ROM, &HE, METHOD_BUFFERED, FILE_READ_ACCESS)
            IOCTL_CDROM_DISK_TYPE = CTL_CODE(FILE_DEVICE_CD_ROM, &H10, METHOD_BUFFERED, FILE_ANY_ACCESS)

       IOCTL_STORAGE_CHECK_VERIFY = CTL_CODE(IOCTL_STORAGE_BASE, &H200, METHOD_BUFFERED, FILE_READ_ACCESS)
      IOCTL_STORAGE_MEDIA_REMOVAL = CTL_CODE(IOCTL_STORAGE_BASE, &H201, METHOD_BUFFERED, FILE_READ_ACCESS)
        IOCTL_STORAGE_EJECT_MEDIA = CTL_CODE(IOCTL_STORAGE_BASE, &H202, METHOD_BUFFERED, FILE_READ_ACCESS)
         IOCTL_STORAGE_LOAD_MEDIA = CTL_CODE(IOCTL_STORAGE_BASE, &H203, METHOD_BUFFERED, FILE_READ_ACCESS)
            IOCTL_STORAGE_RESERVE = CTL_CODE(IOCTL_STORAGE_BASE, &H204, METHOD_BUFFERED, FILE_READ_ACCESS)
            IOCTL_STORAGE_RELEASE = CTL_CODE(IOCTL_STORAGE_BASE, &H205, METHOD_BUFFERED, FILE_READ_ACCESS)
   IOCTL_STORAGE_FIND_NEW_DEVICES = CTL_CODE(IOCTL_STORAGE_BASE, &H206, METHOD_BUFFERED, FILE_READ_ACCESS)
   IOCTL_STORAGE_EJECTION_CONTROL = CTL_CODE(IOCTL_STORAGE_BASE, &H250, METHOD_BUFFERED, FILE_ANY_ACCESS)
        IOCTL_STORAGE_MCN_CONTROL = CTL_CODE(IOCTL_STORAGE_BASE, &H251, METHOD_BUFFERED, FILE_ANY_ACCESS)

     IOCTL_STORAGE_GET_MEDIA_TYPES = CTL_CODE(IOCTL_STORAGE_BASE, &H300, METHOD_BUFFERED, FILE_ANY_ACCESS)
  IOCTL_STORAGE_GET_MEDIA_TYPES_EX = CTL_CODE(IOCTL_STORAGE_BASE, &H301, METHOD_BUFFERED, FILE_ANY_ACCESS)

           IOCTL_STORAGE_RESET_BUS = CTL_CODE(IOCTL_STORAGE_BASE, &H400, METHOD_BUFFERED, FILE_READ_ACCESS)
        IOCTL_STORAGE_RESET_DEVICE = CTL_CODE(IOCTL_STORAGE_BASE, &H401, METHOD_BUFFERED, FILE_READ_ACCESS)
   IOCTL_STORAGE_GET_DEVICE_NUMBER = CTL_CODE(IOCTL_STORAGE_BASE, &H420, METHOD_BUFFERED, FILE_ANY_ACCESS)

     IOCTL_DISK_GET_DRIVE_GEOMETRY = CTL_CODE(IOCTL_DISK_BASE, &H0, METHOD_BUFFERED, FILE_ANY_ACCESS)
     IOCTL_DISK_GET_PARTITION_INFO = CTL_CODE(IOCTL_DISK_BASE, &H1, METHOD_BUFFERED, FILE_READ_ACCESS)
     IOCTL_DISK_SET_PARTITION_INFO = CTL_CODE(IOCTL_DISK_BASE, &H2, METHOD_BUFFERED, FILE_READ_ACCESS Or FILE_WRITE_ACCESS)
       IOCTL_DISK_GET_DRIVE_LAYOUT = CTL_CODE(IOCTL_DISK_BASE, &H3, METHOD_BUFFERED, FILE_READ_ACCESS)
       IOCTL_DISK_SET_DRIVE_LAYOUT = CTL_CODE(IOCTL_DISK_BASE, &H4, METHOD_BUFFERED, FILE_READ_ACCESS Or FILE_WRITE_ACCESS)
                 IOCTL_DISK_VERIFY = CTL_CODE(IOCTL_DISK_BASE, &H5, METHOD_BUFFERED, FILE_ANY_ACCESS)
          IOCTL_DISK_FORMAT_TRACKS = CTL_CODE(IOCTL_DISK_BASE, &H6, METHOD_BUFFERED, FILE_READ_ACCESS Or FILE_WRITE_ACCESS)
        IOCTL_DISK_REASSIGN_BLOCKS = CTL_CODE(IOCTL_DISK_BASE, &H7, METHOD_BUFFERED, FILE_READ_ACCESS Or FILE_WRITE_ACCESS)
            IOCTL_DISK_PERFORMANCE = CTL_CODE(IOCTL_DISK_BASE, &H8, METHOD_BUFFERED, FILE_ANY_ACCESS)
            IOCTL_DISK_IS_WRITABLE = CTL_CODE(IOCTL_DISK_BASE, &H9, METHOD_BUFFERED, FILE_ANY_ACCESS)
                IOCTL_DISK_LOGGING = CTL_CODE(IOCTL_DISK_BASE, &HA, METHOD_BUFFERED, FILE_ANY_ACCESS)
       IOCTL_DISK_FORMAT_TRACKS_EX = CTL_CODE(IOCTL_DISK_BASE, &HB, METHOD_BUFFERED, FILE_READ_ACCESS Or FILE_WRITE_ACCESS)
    IOCTL_DISK_HISTOGRAM_STRUCTURE = CTL_CODE(IOCTL_DISK_BASE, &HC, METHOD_BUFFERED, FILE_ANY_ACCESS)
         IOCTL_DISK_HISTOGRAM_DATA = CTL_CODE(IOCTL_DISK_BASE, &HD, METHOD_BUFFERED, FILE_ANY_ACCESS)
        IOCTL_DISK_HISTOGRAM_RESET = CTL_CODE(IOCTL_DISK_BASE, &HE, METHOD_BUFFERED, FILE_ANY_ACCESS)
      IOCTL_DISK_REQUEST_STRUCTURE = CTL_CODE(IOCTL_DISK_BASE, &HF, METHOD_BUFFERED, FILE_ANY_ACCESS)
           IOCTL_DISK_REQUEST_DATA = CTL_CODE(IOCTL_DISK_BASE, &H10, METHOD_BUFFERED, FILE_ANY_ACCESS)
           IOCTL_DISK_CHECK_VERIFY = CTL_CODE(IOCTL_DISK_BASE, &H200, METHOD_BUFFERED, FILE_READ_ACCESS)
          IOCTL_DISK_MEDIA_REMOVAL = CTL_CODE(IOCTL_DISK_BASE, &H201, METHOD_BUFFERED, FILE_READ_ACCESS)
            IOCTL_DISK_EJECT_MEDIA = CTL_CODE(IOCTL_DISK_BASE, &H202, METHOD_BUFFERED, FILE_READ_ACCESS)
             IOCTL_DISK_LOAD_MEDIA = CTL_CODE(IOCTL_DISK_BASE, &H203, METHOD_BUFFERED, FILE_READ_ACCESS)
                IOCTL_DISK_RESERVE = CTL_CODE(IOCTL_DISK_BASE, &H204, METHOD_BUFFERED, FILE_READ_ACCESS)
                IOCTL_DISK_RELEASE = CTL_CODE(IOCTL_DISK_BASE, &H205, METHOD_BUFFERED, FILE_READ_ACCESS)
       IOCTL_DISK_FIND_NEW_DEVICES = CTL_CODE(IOCTL_DISK_BASE, &H206, METHOD_BUFFERED, FILE_READ_ACCESS)
        IOCTL_DISK_GET_MEDIA_TYPES = CTL_CODE(IOCTL_DISK_BASE, &H300, METHOD_BUFFERED, FILE_ANY_ACCESS)
End Sub

Public Function GetDriveHandle( _
    ByVal drive As String, _
    Optional acc As Long = GENERIC_READ Or GENERIC_WRITE _
) As hFile

    Dim drvchr  As String

    drvchr = Left$(drive, 1)

    GetDriveHandle = FileOpen("\\.\" & drvchr & ":", _
                              acc, _
                              FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                              OPEN_EXISTING)

End Function

Public Function GetBusAddr( _
    ByVal drive As String _
) As SCSI_ADDRESS

    Dim udtAddr As SCSI_ADDRESS
    Dim hDrive  As hFile
    Dim lngRet  As Long
    Dim dwRet   As Long

    hDrive = GetDriveHandle(drive)
    If hDrive.handle = INVALID_HANDLE Then Exit Function

    lngRet = DeviceIoControl(hDrive.handle, _
                             IOCTL_SCSI_GET_ADDRESS, _
                             udtAddr, Len(udtAddr), _
                             udtAddr, Len(udtAddr), _
                             dwRet, ByVal 0&)

    GetBusAddr = udtAddr

    FileClose hDrive
End Function

Private Function CTL_CODE( _
    ByVal lDevType As Long, _
    ByVal lFunction As Long, _
    ByVal lMethod As Long, _
    ByVal lAccess As Long _
) As Long

    CTL_CODE = SHL(lDevType, 16) Or _
               SHL(lAccess, 14) Or _
               SHL(lFunction, 2) Or _
               lMethod
End Function
