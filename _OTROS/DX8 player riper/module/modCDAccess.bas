Attribute VB_Name = "modCDAccess"
Option Explicit

' structs/functions for generic CD/DVD-ROM drafts
'
' some parts translated from AKRip (scsidefs.h)

Public Type SENSE_DATA_FMT
    ErrorCode               As Byte
    SegmentNum              As Byte
    SenseKey                As Byte
    InfoByte0               As Byte
    InfoByte1               As Byte
    InfoByte2               As Byte
    InfoByte3               As Byte
    AddSenLen               As Byte
    ComSpecInfo0            As Byte
    ComSpecInfo1            As Byte
    ComSpecInfo2            As Byte
    ComSpecInfo3            As Byte
    AddSenseCode            As Byte
    AddSenQual              As Byte
    FieldRepUCode           As Byte
    SenKeySpec15            As Byte
    SenKeySpec16            As Byte
    SenKeySpec17            As Byte
    AddSenseBytes           As Byte
End Type

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Enum SENSE_KEYS
    KEY_NOSENSE = &H0         ' No Sense
    KEY_RECERROR = &H1        ' Recovered Error
    KEY_NOTREADY = &H2        ' Not Ready
    KEY_MEDIUMERR = &H3       ' Medium Error
    KEY_HARDERROR = &H4       ' Hardware Error
    KEY_ILLGLREQ = &H5        ' Illegal Request
    KEY_UNITATT = &H6         ' Unit Attention
    KEY_DATAPROT = &H7        ' Data Protect
    KEY_BLANKCHK = &H8        ' Blank Check
    KEY_VENDSPEC = &H9        ' Vendor Specific
    KEY_COPYABORT = &HA       ' Copy Abort
    KEY_EQUAL = &HC           ' Equal (Search)
    KEY_VOLOVRFLW = &HD       ' Volume Overflow
    KEY_MISCOMP = &HE         ' Miscompare (Search)
    KEY_RESERVED = &HF        ' Reserved
End Enum

Public Enum SCSI_STATUS
    STATUS_GOOD = &H0         ' Status Good
    STATUS_CHKCOND = &H2      ' Check Condition
    STATUS_CONDMET = &H4      ' Condition Met
    STATUS_BUSY = &H8         ' Busy
    STATUS_INTERM = &H10      ' Intermediate
    STATUS_INTCDMET = &H14    ' Intermediate-condition met
    STATUS_RESCONF = &H18     ' Reservation conflict
    STATUS_COMTERM = &H22     ' Command Terminated
    STATUS_QFULL = &H28       ' Queue full
    STATUS_TIMEOUT = &HFF     ' Timeout
End Enum

Public Enum SCSI_CMDS
    SCSI_PLAYAUD_10 = &H45    ' Play Audio 10-Byte (O)
    SCSI_PLAYAUD_12 = &HA5    ' Play Audio 12-Byte 12-Byte (O)
    SCSI_PLAYAUDMSF = &H47    ' Play Audio MSF (O)
    SCSI_PLAYA_TKIN = &H48    ' Play Audio Track/Index (O)
    SCSI_PLYTKREL10 = &H49    ' Play Track Relative 10-Byte (O)
    SCSI_PLYTKREL12 = &HA9    ' Play Track Relative 12-Byte (O)
    SCSI_READCDCAP = &H25     ' Read CD-ROM Capacity (MANDATORY)
    SCSI_READHEADER = &H44    ' Read Header (O)
    SCSI_SUBCHANNEL = &H42    ' Read Subchannel (O)
    SCSI_READ_TOC = &H43      ' Read TOC (O)
    SCSI_COMPARE = &H39       ' Compare (O)
    SCSI_FORMAT = &H4         ' Format Unit (MANDATORY)
    SCSI_LCK_UN_CAC = &H36    ' Lock Unlock Cache (O)
    SCSI_PREFETCH = &H34      ' Prefetch (O)
    SCSI_MED_REMOVL = &H1E    ' Prevent/Allow medium Removal (O)
    SCSI_READ6 = &H8          ' Read 6-byte (MANDATORY)
    SCSI_READ10 = &H28        ' Read 10-byte (MANDATORY)
    SCSI_RD_CAPAC = &H25      ' Read Capacity (MANDATORY)
    SCSI_RD_DEFECT = &H37     ' Read Defect Data (O)
    SCSI_READ_LONG = &H3E     ' Read Long (O)
    SCSI_REASS_BLK = &H7      ' Reassign Blocks (O)
    SCSI_RCV_DIAG = &H1C      ' Receive Diagnostic Results (O)
    SCSI_RELEASE = &H17       ' Release Unit (MANDATORY)
    SCSI_REZERO = &H1         ' Rezero Unit (O)
    SCSI_SRCH_DAT_E = &H31    ' Search Data Equal (O)
    SCSI_SRCH_DAT_H = &H30    ' Search Data High (O)
    SCSI_SRCH_DAT_L = &H32    ' Search Data Low (O)
    SCSI_SEEK6 = &HB          ' Seek 6-Byte (O)
    SCSI_SEEK10 = &H2B        ' Seek 10-Byte (O)
    SCSI_SET_LIMIT = &H33     ' Set Limits (O)
    SCSI_START_STP = &H1B     ' Start/Stop Unit (O)
    SCSI_SYNC_CACHE = &H35    ' Synchronize Cache (O)
    SCSI_VERIFY = &H2F        ' Verify (O)
    SCSI_WRITE6 = &HA         ' Write 6-Byte (MANDATORY)
    SCSI_WRITE10 = &H2A       ' Write 10-Byte (MANDATORY)
    SCSI_WRT_VERIFY = &H2E    ' Write and Verify (O)
    SCSI_WRITE_LONG = &H3F    ' Write Long (O)
    SCSI_WRITE_SAME = &H41    ' Write Same (O)
    SCSI_CHANGE_DEF = &H40    ' Change Definition (Optional)
    SCSI_COPY = &H18          ' Copy (O)
    SCSI_COP_VERIFY = &H3A    ' Copy and Verify (O)
    SCSI_INQUIRY = &H12       ' Inquiry (MANDATORY)
    SCSI_LOG_SELECT = &H4C    ' Log Select (O)
    SCSI_LOG_SENSE = &H4D     ' Log Sense (O)
    SCSI_MODE_SEL6 = &H15     ' Mode Select 6-byte (Device Specific)
    SCSI_MODE_SEL10 = &H55    ' Mode Select 10-byte (Device Specific)
    SCSI_MODE_SEN6 = &H1A     ' Mode Sense 6-byte (Device Specific)
    SCSI_MODE_SEN10 = &H5A    ' Mode Sense 10-byte (Device Specific)
    SCSI_READ_BUFF = &H3C     ' Read Buffer (O)
    SCSI_READ_BUFF_CAP = &H5C ' Read Buffer Capacity
    SCSI_SEND_DIAG = &H1D     ' Send Diagnostic (O)
    SCSI_TST_U_RDY = &H0      ' Test Unit Ready (MANDATORY)
    SCSI_WRITE_BUFF = &H3B    ' Write Buffer (O)
    SCSI_BLANK = &HA1         ' Blank
    SCSI_CLOSE_TRK = &H5B     ' Close Track/Session
    SCSI_ERASE10 = &H2C       ' Erase 10-byte
    SCSI_FMT_UNIT = &H4       ' Format Unit
    SCSI_GET_CONF = &H46      ' Get Configuration
    SCSI_GET_EV_ST = &H4A     ' Get Event/Status Notification
    SCSI_GET_PERF = &HAC      ' Get Performance
    SCSI_LOUNLOAD = &HA6      ' Load/Unload
    SCSI_MECH_ST = &HBD       ' Mechanism Status
    SCSI_PAUSE_RESUME = &H4B  ' Pause/Resume
    SCSI_PLAY_CD = &HBC       ' Play CD
    SCSI_READ_CD = &HBE       ' Read CD
    SCSI_READ_CD_MSF = &HB9   ' Read CD MSF
    SCSI_READ_TRK_INF = &H52  ' Read Track Information
    SCSI_REQ_SENSE = &H3      ' Request Sense
    SCSI_SET_SPEED = &HBB     ' Set CD Speed
End Enum

Public Const MAXLUN         As Long = 8     ' Maximum Logical Unit Id
Public Const MAXTARG        As Long = 8     ' Maximum Target Id
Public Const MAX_SCSI_LUNS  As Long = 64    ' Maximum Number of SCSI LUNs
Public Const MAX_NUM_HA     As Long = 8     ' Maximum Number of SCSI HA's

' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Type MSF
    m                       As Byte
    s                       As Byte
    F                       As Byte
End Type

Public Type SCSI_INQUIRY
    qualifier               As Byte
    rsvd1                   As Byte
    Version                 As Byte
    respfmt                 As Byte
    addlen                  As Byte
    rsvd2                   As Byte
    stuff(1)                As Byte
    vendor(7)               As Byte
    product(15)             As Byte
    revision(3)             As Byte
    rsvd3(1)                As Byte
    stuff2(37)              As Byte
End Type

Public Function MSF2STR( _
    fmt As MSF _
) As String

    MSF2STR = format(fmt.m, "00") & ":" & _
              format(fmt.s, "00") & ":" & _
              format(fmt.F, "00")
End Function

Public Function MSF2LBA( _
    fmt As MSF, _
    Optional pos As Boolean _
) As Long

    With fmt
        MSF2LBA = CLng(.m) * 60 * 75 + (.s * 75) + .F
    End With

    If fmt.m < 90 Or pos Then
        MSF2LBA = MSF2LBA - 150
    Else
        MSF2LBA = MSF2LBA - 450150
    End If
End Function

Public Function LBA2MSF( _
    ByVal LBA As Long _
) As MSF

    Dim m As Long, s As Long, F As Long, Start As Long

    Start = Choose(Abs(CBool(LBA >= -150)) + 1, 450150, 150)

    With LBA2MSF
        .m = Fix((LBA + Start) / (60& * 75&))
        .s = Fix((LBA + Start - .m * 60& * 75&) / 75&)
        .F = Fix(LBA + Start - .m * 60& * 75& - .s * 75&)
    End With
End Function

Public Function VarToLBA( _
    ParamArray fmt() As Variant _
) As Long

    Dim btLng()     As Byte
    Dim lng         As Long

    If TypeName(fmt(0)) = "String" Then
        If Len(fmt(0)) = 4 Then
            btLng = StrConv(fmt(0), vbFromUnicode)
            DXCopyMemory lng, btLng(0), 4
            VarToLBA = lng
        Else
            VarToLBA = Val(fmt(0))
        End If
    ElseIf UBound(fmt) = 3 Then
        ReDim btLng(3) As Byte
        btLng(0) = fmt(0)
        btLng(1) = fmt(1)
        btLng(2) = fmt(2)
        btLng(3) = fmt(3)
        DXCopyMemory lng, btLng(0), 4
        VarToLBA = lng
    ElseIf UBound(fmt(0)) = 3 Then
        ReDim btLng(3) As Byte
        btLng(0) = fmt(0)(0)
        btLng(1) = fmt(0)(1)
        btLng(2) = fmt(0)(2)
        btLng(3) = fmt(0)(3)
        DXCopyMemory lng, btLng(0), 4
        VarToLBA = lng
    End If
End Function

Public Function VarToMSF( _
    ParamArray fmt() As Variant _
) As MSF

    If TypeName(fmt(0)) = "String" Then
        VarToMSF.m = Left$(fmt(0), InStr(fmt(0), ":") - 1)
        fmt(0) = Mid$(fmt(0), InStr(fmt(0), ":") + 1)
        VarToMSF.s = Left$(fmt(0), InStr(fmt(0), ":") - 1)
        fmt(0) = Val(Mid$(fmt(0), InStr(fmt(0), ":") + 1))
        VarToMSF.F = fmt(0)
    ElseIf UBound(fmt) = 2 Then
        VarToMSF.m = fmt(0)
        VarToMSF.s = fmt(1)
        VarToMSF.F = fmt(2)
    ElseIf UBound(fmt) = 3 Then
        VarToMSF.m = fmt(1)
        VarToMSF.s = fmt(2)
        VarToMSF.F = fmt(3)
    ElseIf UBound(fmt(0)) = 2 Then
        VarToMSF.m = fmt(0)(0)
        VarToMSF.s = fmt(0)(1)
        VarToMSF.F = fmt(0)(2)
    ElseIf UBound(fmt(0)) = 3 Then
        VarToMSF.m = fmt(0)(1)
        VarToMSF.s = fmt(0)(2)
        VarToMSF.F = fmt(0)(3)
    End If
End Function
