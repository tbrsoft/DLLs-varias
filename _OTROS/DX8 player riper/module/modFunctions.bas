Attribute VB_Name = "modFunctions"
Option Explicit

' misc functions

Public Declare Function GetVersionEx Lib "kernel32" _
Alias "GetVersionExA" ( _
    lpVersionInformation As OSVERSIONINFO _
) As Long

Public Declare Function GetDriveType Lib "kernel32" _
Alias "GetDriveTypeA" ( _
    ByVal nDrive As String _
) As Long

Public Declare Function GlobalUnlock Lib "kernel32.dll" ( _
    ByVal hMem As Long _
) As Long

Public Declare Function GlobalAlloc Lib "kernel32.dll" ( _
    ByVal wFlags As GMEMFlags, _
    ByVal dwBytes As Long _
) As Long

Public Declare Function GlobalFree Lib "kernel32.dll" ( _
    ByVal hMem As Long _
) As Long

Public Declare Function GlobalLock Lib "kernel32.dll" ( _
    ByVal hMem As Long _
) As Long

Public Enum GMEMFlags
    GMEM_FIXED = &H0
    GMEM_MOVEABLE = &H2
    GMEM_ZEROINIT = &H40
End Enum

Public Type OSVERSIONINFO
  dwOSVersionInfoSize   As Long
  dwMajorVersion        As Long
  dwMinorVersion        As Long
  dwBuildNumber         As Long
  dwPlatformId          As Long
  szCSDVersion          As String * 128
End Type

Public Type MEMORY
    handle              As Long
    pointer             As Long
    bytes               As Long
End Type

Public Type LARGE_INTEGER
    lo                  As Long
    hi                  As Long
End Type

Public Const VER_PLATFORM_WIN32s        As Long = 0
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT      As Long = 2

Private lngPower2(31)   As Long

Public Function AllocateMem( _
    ByVal bytes As Long _
) As MEMORY

    With AllocateMem
        .bytes = bytes
        .handle = GlobalAlloc(GMEM_FIXED Or GMEM_ZEROINIT, bytes)
        .pointer = GlobalLock(.handle)
    End With
End Function

Public Sub FreeMem( _
    mem As MEMORY _
)

    With mem
        GlobalUnlock .handle
        GlobalFree .handle

        ZeroMemory mem, Len(mem)
    End With
End Sub

Public Function GetWindowsVersion() As OSVERSIONINFO
    Dim udtVersion  As OSVERSIONINFO

    udtVersion.dwOSVersionInfoSize = Len(udtVersion)
    GetVersionEx udtVersion

    GetWindowsVersion = udtVersion
End Function

Public Function GetDirLevel( _
    ByVal strFile As String _
) As Integer

    GetDirLevel = UBound(Split(strFile, "\"))
End Function

Public Function GetExtension( _
    ByVal strFile As String _
) As String

    If InStr(strFile, ".") > 0 Then
        GetExtension = Mid$(strFile, InStrRev(strFile, ".") + 1)
    Else
        GetExtension = strFile
    End If
End Function

Public Function FmtStrLen( _
    ByVal B As String, _
    Length As Integer _
) As String

    B = String(Length - Len(B), "0") & B
    FmtStrLen = B
End Function

Public Function RemNullChars( _
    ByVal strString As String _
) As String

    If InStr(strString, Chr$(0)) > 0 Then
        RemNullChars = Left$(strString, InStr(strString, Chr$(0)) - 1)
    Else
        RemNullChars = strString
    End If

End Function

Public Function IsBitSet( _
    ByVal value As Long, _
    bit As Byte _
) As Boolean

    IsBitSet = CBool(value And 2 ^ bit)
End Function

Public Function MKWord( _
    ByVal Bh As Byte, _
    ByVal Bl As Byte _
) As Integer

    MKWord = SHL(Bh, 8) Or Bl
End Function

Public Function MKDWord( _
    ByVal Wh As Integer, _
    ByVal Wl As Integer _
) As Long

    MKDWord = SHL(Wh, 16) Or Wl
End Function

Public Function LoWord( _
    ByVal DWord As Long _
) As Long

  LoWord = DWord And &HFFFF&
End Function

Public Function HiWord( _
    ByVal DWord As Long _
) As Long

  HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function LoByte( _
    ByVal Word As Integer _
) As Byte

  LoByte = Word And &HFF
End Function

Public Function HiByte( _
    ByVal Word As Integer _
) As Byte

  HiByte = (Word And &HFF00&) \ &H100
End Function

Public Function LoNibble( _
    ByVal Bt As Byte _
) As Byte

    LoNibble = Bt And &HF
End Function

Public Function HiNibble( _
    ByVal Bt As Byte _
) As Byte

    HiNibble = (Bt And &HF0) \ &H10
End Function

' >> Operator
' by VB-Accelerator
Public Function SHR( _
    ByVal lThis As Long, _
    ByVal lBits As Long _
) As Long

    Static Init As Boolean

    If Not Init Then InitShifting: Init = True

    If (lBits <= 0) Then
        SHR = lThis
    ElseIf (lBits > 63) Then
        Exit Function
    ElseIf (lBits > 31) Then
        SHR = 0
    Else
        If (lThis And lngPower2(31)) = lngPower2(31) Then
            SHR = (lThis And &H7FFFFFFF) \ lngPower2(lBits) Or lngPower2(31 - lBits)
        Else
            SHR = lThis \ lngPower2(lBits)
        End If
    End If

End Function

' << Operator
' by VB-Accelerator
Public Function SHL( _
    ByVal lThis As Long, _
    ByVal lBits As Long _
) As Long

    Static Init As Boolean

    If Not Init Then InitShifting: Init = True

    If (lBits <= 0) Then
        SHL = lThis
    ElseIf (lBits > 63) Then
        Exit Function
    ElseIf (lBits > 31) Then
        SHL = 0
    Else
        If (lThis And lngPower2(31 - lBits)) = lngPower2(31 - lBits) Then
            SHL = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits) Or lngPower2(31)
        Else
            SHL = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits)
        End If
    End If

End Function

' power of 2
Private Sub InitShifting()
    Dim i   As Long
    For i = 0 To 30: lngPower2(i) = 2& ^ i: Next
    lngPower2(31) = &H80000000
End Sub

Public Function Dbl2LargeInt( _
    ByVal dbl As Double _
) As LARGE_INTEGER

    Dim strHex  As String
    Dim qwRet   As LARGE_INTEGER

    strHex = FmtStrLen(DecToHex(dbl), 16)

    ' lo DWORD
    qwRet.lo = Val("&H" & Right$(strHex, 8) & "&")
    ' hi DWORD
    qwRet.hi = Val("&H" & Left$(strHex, 8) & "&")

    Dbl2LargeInt = qwRet
End Function

' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=49987&lngWId=1
Public Function DecToHex( _
    ByVal nSource As Double _
) As String

    Const BASECHAR  As String = "0123456789ABCDEF"
    Dim n           As Double
    Dim nSrc        As Double

    On Error GoTo ErrorLine

    If (nSource = 0) Then
        DecToHex = "00"
        Exit Function
    End If

    If (nSource < 2147483648#) Then
        DecToHex = Hex(nSource)
    Else
        nSrc = nSource

        Do
            n = CDec(nSrc - (16 * Int(nSrc / 16)))
            DecToHex = Mid$(BASECHAR, n + 1, 1) & DecToHex
            nSrc = CDec(Int(nSrc / 16))
        Loop While (nSrc > 0)

    End If

    If (Len(DecToHex) Mod 2) Then
        DecToHex = "0" & DecToHex
    End If
    Exit Function

ErrorLine:
    DecToHex = ""
End Function
