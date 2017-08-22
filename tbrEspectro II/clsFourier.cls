VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFourier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fast Fourier Transformation - FFT
'
' ...and it is fast indeed - will transform 2048 samples in
' S1 mSec now on an Atlon 1800 MHz CPU with a clockrate of
' 1533 MHz (with a little help from my friends)
'
' by Ulli

Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As Currency _
) As Long

Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency _
) As Long

Private Declare Sub MemCopy Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long _
)

Private AtStart             As Currency
Private AtEnd               As Currency
Private CPUSpeed            As Currency

Private Const Mirror_FON    As Long = 11
Private Mirror_Bin()        As Byte
Private Const Mirror_Hex    As String = "31 C0 8B 54 24 08 8B 4C 24 0C 67 E3 07 D1 DA 11 C0 49 75 F9 8B 54 24 10 89 02 C2 10 00"

Private Const Ioor          As String = "Index out of range or number of samples unknown"
Private Const Nosm          As String = "Number of samples must be a positive value of the form 2 ^ n"

Private UnknownSize         As Boolean
Private NeedsDoing          As Boolean
Private myTimeWindow        As Long
Private UBSamples           As Long     'upper bound of samples
Private NumBits             As Long     'number of bits needed to express above
Private StageSize           As Long     'the number of samples in current computation stage
Private NumButter           As Long     'the number of butterflies in current stage
Private i                   As Long     'helpers, enumerators and such
Private j                   As Long
'Private Align8              As Long     'this padding (if necessary) improves speed by about 20 to 25%
'                                       'apparently there's a severe penalty on misaliged operands (at least with the Athlon)
Private Pi                  As Double   'what it says:  pi
Private TwoPi               As Double
Private tmp                 As Double

Private Type Sample                     'sample consists of a real and an imaginary value in gaussian complex plane
    Real                    As Double
    Imag                    As Double
End Type

Private s                   As Sample
Private t                   As Sample
Private U                   As Sample
Private Values()            As Sample

Private Sub Butterfly( _
    ByRef pS As Sample, _
    ByRef pU As Sample, _
    ByRef oJ As Sample, _
    ByRef oK As Sample _
)

    t.Real = pU.Real * oK.Real - pU.Imag * oK.Imag
    t.Imag = pU.Imag * oK.Real + pU.Real * oK.Imag
    oK.Real = oJ.Real - t.Real
    oJ.Real = oJ.Real + t.Real
    oK.Imag = oJ.Imag - t.Imag
    oJ.Imag = oJ.Imag + t.Imag
    tmp = pS.Real * pU.Real + pS.Imag * pU.Imag
    pU.Imag = pU.Imag + pS.Imag * pU.Real - pS.Real * pU.Imag
    pU.Real = pU.Real - tmp
End Sub

Private Sub Class_Initialize()
    Pi = 4 * Atn(1)
    TwoPi = Pi + Pi
    ReDim Values(0)

    UnknownSize = True
    QueryPerformanceFrequency CPUSpeed

    Patch Mirror_FON, Mirror_Hex, Mirror_Bin

    i = VarPtr(Pi) Mod 8

'    If i Then
'        MsgBox "For improved speed insert padding of " & 8 - i & " bytes before Variable Pi", vbExclamation, "Developer:"
'    End If
End Sub

Public Property Get ComplexOut( _
    Index As Long _
) As Double

    With GetIt(Index)
        ComplexOut = Sqr(.Real * .Real + .Imag * .Imag)
    End With 'GETIT(INDEX)
End Property

Private Function GetIt( _
    Index As Long _
) As Sample

    If UnknownSize Or Index < 1 Or Index > UBSamples + 1 Then
        Err.Raise 381, , Ioor
      Else 'NOT UNKNOWNSIZE...
        If NeedsDoing Then
            NeedsDoing = False

            '=======================================================================================
            'Begin Fast Fourier Transformation

            QueryPerformanceCounter AtStart

            StageSize = 1
            Do
                'divide and conquer
                NumButter = StageSize
                StageSize = NumButter * 2
                tmp = Pi / StageSize
                s.Real = Sin(tmp)
                s.Real = 2 * s.Real * s.Real
                s.Imag = Sin(tmp * 2)
                For i = 0 To UBSamples Step StageSize
                    U.Real = 1
                    U.Imag = 0
                    For j = i To i + NumButter - 1
                        Butterfly s, U, Values(j), Values(j + NumButter) 'butterfly calculation
                Next j, i
            Loop Until StageSize > UBSamples

            QueryPerformanceCounter AtEnd

            'End Fast Fourier Transformation
            '=======================================================================================

        End If
        GetIt = Values(Index - 1)
    End If
End Function

Public Property Let ImagIn( _
    ByVal Index As Long, _
    nuValueIn As Double _
)

    If UnknownSize Or Index < 1 Or Index > UBSamples + 1 Then
        Err.Raise 381, , Ioor
      Else 'NOT UNKNOWNSIZE...
        Values(Mirror(Index - 1, NumBits)).Imag = nuValueIn
        NeedsDoing = True
    End If
End Property

Public Property Get ImagOut( _
    ByVal Index As Long _
) As Double

    ImagOut = GetIt(Index).Imag
End Property

Private Function Mirror( _
    ByVal Index As Long, _
    ByVal NumBits As Long _
) As Long

  'dummy -- will be patched by
  '
  '         xor   eax, eax          ;clear result
  '         mov   edx, [esp + 8]    ;get Index from stack
  '         mov   ecx, [esp + 12]   ;get NumBits from stack
  '         jcxz  GetOut            ;zero? -> get out
  '     Again:
  '         rcr   edx, 1            ;shift LSB of index into carry flag
  '         adc   eax, eax          ;double result and add carry
  '         dec   ecx               ;NumBits - 1
  '         jnz   Again             ;not zero? -> do it again
  '     GetOut:
  '         mov   edx, [esp + 16]   ;get address of function variable from stack
  '         mov   [edx], eax        ;put result there
  '         ret   16                ;return tidying stack

    For j = 1 To NumBits
        Mirror = Mirror * 2 Or (Index And 1)
        Index = Index \ 2
    Next j
End Function

Public Property Let NumberOfSamples( _
    ByVal nuNumSam As Long _
)

    If nuNumSam > 1 And (nuNumSam - 1 And nuNumSam) = 0 Then
        ReDim Values(0 To nuNumSam - 1)
        UnknownSize = False
        UBSamples = nuNumSam - 1
        NumBits = Log(nuNumSam) / Log(2) 'the number of bits needed to express UBSamples
        NeedsDoing = True
      Else 'NOT NUNUMSAM...
        Err.Raise 380, , Nosm
    End If
End Property

Private Sub Patch( _
    ByVal FON As Long, _
    HexCode As String, _
    BinCode() As Byte _
) 'FON is the Function's Ordinal Number in vTable

  'Convert hex to binary and patch vTable entry

  Dim st()          As String
  Dim p             As Long
  Dim CodeAddress   As Long
  Dim VTableAddress As Long

    st = Split(HexCode, " ")
    ReDim BinCode(0 To UBound(st))
    'Convert hex to binary
    For p = 0 To UBound(st)
        BinCode(p) = Val("&H" & st(p))
    Next p

    'Patch vTable entry
    CodeAddress = VarPtr(BinCode(0))
    MemCopy VarPtr(VTableAddress), ObjPtr(Me), 4 'get vTable address
    MemCopy VTableAddress + FON * 4 + 28, VarPtr(CodeAddress), 4 'patch proper entry in vTable
End Sub

Public Property Let RealIn( _
    ByVal Index As Long, _
    nuValueIn As Double _
)

    If UnknownSize Or Index < 1 Or Index > UBSamples + 1 Then
        Err.Raise 381, , Ioor
      Else 'NOT UNKNOWNSIZE...
        Index = Index - 1
        With Values(Mirror(Index, NumBits))
            If myTimeWindow Then
                tmp = TwoPi * Index / UBSamples
                'three term blackman time window function
                .Real = nuValueIn * (0.42 - 0.5 * Cos(tmp) + 0.08 * Cos(2 * tmp))
              Else 'MYTIMEWINDOW = FALSE/0
                .Real = nuValueIn / 2
            End If
            .Imag = 0
        End With 'VALUES(MIRROR(INDEX,
        NeedsDoing = True
    End If
End Property

Public Property Get RealOut( _
    ByVal Index As Long _
) As Double

    RealOut = GetIt(Index).Real
End Property

Public Property Get Timing( _
) As Single

    Timing = (AtEnd - AtStart) / CPUSpeed * 1000
End Property

Public Property Let WithTimeWindow( _
    ByVal nuTimeWindow As Long _
)

    myTimeWindow = nuTimeWindow
End Property

':) Ulli's VB Code Formatter V2.21.6 (2006-Apr-24 11:50)  Decl: 45  Code: 201  Total: 246 Lines
':) CommentOnly: 31 (12,6%)  Commented: 19 (7,7%)  Empty: 57 (23,2%)  Max Logic Depth: 6
