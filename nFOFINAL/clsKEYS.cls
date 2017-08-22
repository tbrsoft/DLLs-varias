VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKEYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mWinVersion As String

Public mCountry As String

Public mBiosDate As String ' solo 98, no permitido en XP
Public mBIOSCaption As String 'solo con WMI (no 98)
Public mBIOSDescription As String 'solo con WMI (no 98)
Public mBIOSIdentificationCode As String 'solo con WMI (no 98)
Public mBIOSManufacturer As String 'solo con WMI (no 98)
Public mBIOSName As String 'solo con WMI (no 98)
Public mBIOSReleaseDate As String 'solo con WMI (no 98)
Public mBIOSSerialNumber As String 'solo con WMI (no 98)
Public mBIOSSMBIOSBIOSVersion As String 'solo con WMI (no 98)
Public mBIOSSMBIOSMajorVersion As String 'solo con WMI (no 98)
Public mBIOSSMBIOSMinorVersion As String 'solo con WMI (no 98)
Public mBIOSVersion As String 'solo con WMI (no 98)

Public mPROCInfoReserved As String
Public mPROCType As String 'solo con WMI (no 98)
Public mPROCRevis As String 'solo con WMI (no 98)
Public mPROCId As String 'solo con WMI (no 98)
Public mPROCCaption As String 'solo con WMI (no 98)
Public mPROCCCspeed As String 'solo con WMI (no 98)
Public mPROCDeviceId As String 'solo con WMI (no 98)
Public mPROCFamily As String 'solo con WMI (no 98)

Private Declare Function GetLocaleInfo Lib "kernel32" Alias _
    "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Declare Sub GetMem1 Lib "msvbvm60.dll" (ByVal _
   MemAddress As Long, var As Byte)

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

'para el pais e idioma
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SENGCOUNTRY = &H1002 '  English name of country
Private Const LOCALE_SENGLANGUAGE = &H1001  '  English name of language
Private Const LOCALE_SNATIVELANGNAME = &H4  '  native name of language
Private Const LOCALE_SNATIVECTRYNAME = &H8  '  native name of country

Private VW As New clsWindowsVERSION

Public Sub LeerTodo()
    Terr.Anotar "uaa"
    'PR.ShowProgreso "Reserved", 25
    '---------------------------------------------------------
    'RESERVED
    '---------------------------------------------------------
    Dim RSV As String
    
    Dim INFO As SYSTEM_INFO
    
    GetSystemInfo INFO
    
    Dim GUIDtmp As String 'no es guid, es un valor unico para cada PC
    'este reserved es un numero entre 50.000.000 y 140.000.000
    RSV = CStr(INFO.dwReserved)
    Terr.Anotar "uab", RSV
    If Len(RSV) < 3 Then
        'no es compatible en esta PC
        'meter un valor fijo
        RSV = "FF"
    Else
        'pasar a hex para que no sea igual a antes
        RSV = Hex(CLng(RSV))
    End If
    Terr.Anotar "uac", RSV
    mPROCInfoReserved = RSV
    
    'PR.ShowProgreso "Windows", 30
    '---------------------------------------------------------
    'Windows Version
    '---------------------------------------------------------
    mWinVersion = VW.GetVersionString
    Terr.Anotar "uad", mWinVersion
    
    'pr.ShowProgreso "Bios98", 35
    '---------------------------------------------------------
    'Bios 98
    '---------------------------------------------------------
    Dim sBios As String
    ' el getmem se que anda en 98 y me (no comprobado en 95s)
    If VW.Version = Win98 Or VW.Version = Win98SE Or VW.Version = WinME Then
        'EN XP DA ERROR ya que no se puede acceder asi a la memoria!!!
        Dim P As Byte, MemAddr As Long
        Dim i As Integer
        'start of bios serial number ?&HFE0C0
        MemAddr = &HFE000
        Terr.Anotar "uae"
        For i = 0 To 331
            Call GetMem1(MemAddr + i, P)
            'get printable characters
            If P > 31 And P <= 128 Then sBios = sBios & Chr$(P)
        Next i
        mBiosDate = sBios
    Else
        Terr.Anotar "uaf"
        mBiosDate = "noX11"
    End If
    'pr.ShowProgreso "Bios WMI", 40
    '---------------------------------------------------------
    'Bios con WMI (no en 98)
    '---------------------------------------------------------
    
    ' si es XP el Get Mem se va a la bosta y se clava la maquina sin error !!!!
    If VW.Version = Win2000 Or VW.Version = WinXp Or _
        VW.Version = WinXP2 Or VW.Version = WinNT4 Then
    
        On Error GoTo NoWBmem
        Terr.Anotar "uag"
        Dim ObjSet As SWbemObjectSet
        Dim SERV As SWbemServices
        Set SERV = GetObject("WinMgmts:")
        Set ObjSet = Nothing
        Set ObjSet = SERV.InstancesOf("Win32_Bios")
        If ObjSet.Count = 1 Then
            For Each BIOS In ObjSet
                mBIOSCaption = CStr(NN(BIOS.Caption))
                mBIOSDescription = CStr(NN(BIOS.Description))
                mBIOSIdentificationCode = CStr(NN(BIOS.IdentificationCode))
                mBIOSManufacturer = CStr(NN(BIOS.Manufacturer))
                mBIOSName = CStr(NN(BIOS.Name))
                mBIOSReleaseDate = CStr(NN(BIOS.ReleaseDate))
                mBIOSSerialNumber = CStr(NN(BIOS.SerialNumber))
                mBIOSSMBIOSBIOSVersion = CStr(NN(BIOS.SMBIOSBIOSVersion))
                mBIOSSMBIOSMajorVersion = CStr(NN(BIOS.SMBIOSMajorVersion))
                mBIOSSMBIOSMinorVersion = CStr(NN(BIOS.SMBIOSMinorVersion))
                mBIOSVersion = CStr(NN(BIOS.Version))
            Next
        End If
    End If
    Set SERV = Nothing: Set ObjSet = Nothing
    GoTo SIG01
            
NoWBmem:
        mBIOSCaption = "noX11"
        mBIOSDescription = "noX11"
        mBIOSIdentificationCode = "noX11"
        mBIOSManufacturer = "noX11"
        mBIOSName = "noX11"
        mBIOSReleaseDate = "noX11"
        mBIOSSerialNumber = "noX11"
        mBIOSSMBIOSBIOSVersion = "noX11"
        mBIOSSMBIOSMajorVersion = "noX11"
        mBIOSSMBIOSMinorVersion = "noX11"
        mBIOSVersion = "noX11"
SIG01:
    'pr.ShowProgreso "Proc", 50
    '---------------------------------------------------------
    'PROCESADOR
    '---------------------------------------------------------
    On Error GoTo NoWBmem2
    
    Dim ObjSet2 As SWbemObjectSet
    Dim SERV2 As SWbemServices
    Set SERV2 = GetObject("WinMgmts:")
    Set ObjSet2 = Nothing
    Set ObjSet2 = SERV2.InstancesOf("Win32_Processor")
    If ObjSet2.Count = 1 Then
        For Each MICRO In ObjSet2
            mPROCType = CStr(NN(MICRO.ProcessorType))
            mPROCRevis = CStr(NN(MICRO.Revision))
            mPROCId = CStr(NN(MICRO.ProcessorId))
            mPROCCaption = CStr(NN(MICRO.Caption))
            mPROCCCspeed = CStr(NN(MICRO.CurrentClockSpeed))
            mPROCDeviceId = CStr(NN(MICRO.DeviceID))
            mPROCFamily = CStr(NN(MICRO.Family))
        Next
    End If
    
    
    Set SERV2 = Nothing: Set ObjSet2 = Nothing
    GoTo SIG02
        
NoWBmem2:
    mPROCType = "noX11"
    mPROCRevis = "noX11"
    mPROCId = "noX11"
    mPROCCaption = "noX11"
    mPROCCCspeed = "noX11"
    mPROCDeviceId = "noX11"
    mPROCFamily = "noX11"
    Terr.Anotar "uah"
SIG02:
    'pr.ShowProgreso "Pais", 80
    '---------------------------------------------------------
    'PAIS
    '---------------------------------------------------------
    Dim TMP As String
    TMP = GetInfo(LOCALE_SENGCOUNTRY) ' + "(" + GetInfo(LOCALE_SNATIVECTRYNAME) & ")"
    
    mCountry = TMP
    Terr.Anotar "uai", mCountry
End Sub

Private Function LeerArch1Linea(Arch As String) As String
    'If Dir(Arch) = "" Then
    If FSO.FileExists(Arch) = False Then
        LeerArch1Linea = "No existe archivo"
        Exit Function
    End If
    'Dim Libre As Long, TMP As String
    'Libre = FreeFile
    'Open Arch For Input As Libre
    Set TE = FSO.OpenTextFile(Arch, ForReading, False)
        'Line Input #Libre, TMP
    LeerArch1Linea = TE.ReadLine
    'Close Libre
    TE.Close
    'LeerArch1Linea = TMP
End Function

Private Function HEXtoLONG(n As String)
    'recibe el hex en str y devuelve un numero en str
    
    Dim Letra As String
    Dim c As Long
    Dim NumeroActual As Long
    Dim ACUM ' As Double
    For c = 1 To Len(n)
        Letra = Mid(n, c, 1)
        Select Case Letra
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                NumeroActual = Val(Letra)
            Case "A"
                NumeroActual = 10
            Case "B"
                NumeroActual = 11
            Case "C"
                NumeroActual = 12
            Case "D"
                NumeroActual = 13
            Case "E"
                NumeroActual = 14
            Case "F"
                NumeroActual = 15
        End Select
        Dim ToSum ' As Double
        ToSum = NumeroActual * (16 ^ (Len(n) - c))
        ACUM = ACUM + ToSum
        'Label10 = Label10 + "LETRA: " + Letra + "=" + CStr(ToSum) + vbCrLf
        
    Next
    
    HEXtoLONG = CStr(ACUM)
End Function

Private Function NN(Val, Optional DEfault = "NULO")
    'No Nulo
    If IsNull(Val) Then
        NN = DEfault
    Else
        NN = Val
    End If
End Function

Private Function GetInfo(ByVal lInfo As Long) As String
    Dim Buffer As String, Ret As String
    Buffer = String$(256, 0)
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    If Ret > 0 Then
        GetInfo = Left$(Buffer, Ret - 1)
    Else
        GetInfo = ""
    End If
End Function

Private Sub Class_Initialize()
    VW.GetVersion
End Sub
