VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCMPreparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' convert data using to PCM using the ACM,
' if supported by it
Private Const OUTPUT_BUFFER_MS  As Long = 500

Private btwfxin()       As Byte
Private wfxinp          As MMWAVEFORMATEX
Private wfxout          As MMWAVEFORMATEX

Private hStream         As HACMSTREAM

Private btInput()       As Byte
Private intOutput()     As Integer

Private lngInputLen     As Long
Private lngOutputLen    As Long

Private lngPosInBuffer  As Long
Private lngBufferData   As Long

Private blnEndOfStream  As Boolean
Private blnFirst        As Boolean

' stream to obtain data from
Private clsStream       As ISoundStream
' DSP Callback
Private clsCB           As IDSPCallback

' bytes to carry after a conversion
Private lngKeepInBuffer As Long

Public Property Get EndOfStream() As Boolean
    EndOfStream = blnEndOfStream
End Property

Public Property Get OutputSamplerate() As Long
    OutputSamplerate = wfxout.nSamplesPerSec
End Property

Public Property Get OutputChannels() As Integer
    OutputChannels = wfxout.nChannels
End Property

Public Property Get OutputBitsPerSample() As Integer
    OutputBitsPerSample = wfxout.wBitsPerSample
End Property

Private Function Convert( _
    ByVal lngInLen As Long, _
    ByVal lngOutLen As Long, _
    lngInUsed As Long, _
    lngOutUsed As Long, _
    Optional ByVal LastConversion As Boolean = False _
) As Boolean

    Dim lngFlags    As Long
    Dim udtHdr      As ACMSTREAMHEADER

    lngFlags = ACM_STREAMCONVERTF_BLOCKALIGN

    If blnFirst Then _
        lngFlags = lngFlags Or ACM_STREAMCONVERTF_START
    If LastConversion Then
        lngFlags = lngFlags Or ACM_STREAMCONVERTF_END
    End If

    With udtHdr
        .cbStruct = Len(udtHdr)
        .cbSrcLength = lngInLen
        .cbDstLength = lngOutLen
        .pbDst = VarPtr(intOutput(0))
        .pbSrc = VarPtr(btInput(0))
    End With

    acmStreamPrepareHeader hStream, udtHdr, 0

    blnFirst = False

    If 0 = acmStreamConvert(hStream, udtHdr, lngFlags) Then
        lngInUsed = udtHdr.cbSrcLengthUsed
        lngOutUsed = udtHdr.cbDstLengthUsed

        lngKeepInBuffer = udtHdr.cbSrcLength - udtHdr.cbSrcLengthUsed
        If lngKeepInBuffer > 0 Then
            DXCopyMemory btInput(0), btInput(lngInLen - lngKeepInBuffer), lngKeepInBuffer
        End If

        Convert = True
    End If

    acmStreamUnprepareHeader hStream, udtHdr, 0
End Function

Private Function FillBuffer() As Boolean
    Dim lngRead     As Long
    Dim lngWritten  As Long
    Dim udeRet      As STREAM_ERR

    If blnEndOfStream Then
        lngBufferData = 0
        lngPosInBuffer = 0
        ZeroMemory intOutput(0), lngOutputLen
        ZeroMemory btInput(0), lngInputLen
        Exit Function
    End If

    udeRet = clsStream.Read(VarPtr(btInput(lngKeepInBuffer)), lngInputLen - lngKeepInBuffer, lngRead)

    If udeRet <> STREAM_OK Then
        blnEndOfStream = True
    End If

    Convert lngRead + lngKeepInBuffer, lngOutputLen, _
            lngRead, lngWritten, _
            blnEndOfStream

    lngPosInBuffer = 0
    lngBufferData = lngWritten

    DoDSP

    FillBuffer = True
End Function

Private Sub DoDSP()
    If lngBufferData = 0 Then Exit Sub
    If OutputBitsPerSample < 16 Then Exit Sub

    clsCB.Samples intOutput, lngBufferData, OutputChannels
End Sub

Public Function GetSamples( _
    ByVal pBuffer As Long, _
    ByVal cBytes As Long, _
    cBytesReturned As Long _
) As Boolean

    If clsStream Is Nothing Then Exit Function
    GetSamples = True

    cBytesReturned = 0

    Do While cBytesReturned < cBytes

        ' PCM buffer empty
        ' either beginning or end of the stream
        If lngBufferData = 0 Then
            If Not FillBuffer Then Exit Function

        ' not enough data to fill the requested buffer
        ElseIf (lngBufferData - lngPosInBuffer) < (cBytes - cBytesReturned) Then
            If 0 < (lngBufferData - lngPosInBuffer) Then

                If 0 = IsBadReadPtr(ByVal VarPtr(intOutput(0)) + lngPosInBuffer, lngBufferData - lngPosInBuffer) Then
                    If 0 = IsBadWritePtr(ByVal pBuffer + cBytesReturned, lngBufferData - lngPosInBuffer) Then
                        DXCopyMemory ByVal pBuffer + cBytesReturned, ByVal VarPtr(intOutput(0)) + lngPosInBuffer, lngBufferData - lngPosInBuffer
                    End If
                End If

                cBytesReturned = cBytesReturned + (lngBufferData - lngPosInBuffer)
            End If

            If Not FillBuffer Then Exit Function

        ' enough data to fill the requested buffer
        Else
            If 0 = IsBadReadPtr(ByVal VarPtr(intOutput(0)) + lngPosInBuffer, cBytes - cBytesReturned) Then
                If 0 = IsBadWritePtr(ByVal pBuffer + cBytesReturned, cBytes - cBytesReturned) Then
                    DXCopyMemory ByVal pBuffer + cBytesReturned, ByVal VarPtr(intOutput(0)) + lngPosInBuffer, cBytes - cBytesReturned
                End If
            End If

            lngPosInBuffer = lngPosInBuffer + (cBytes - cBytesReturned)
            cBytesReturned = cBytesReturned + (cBytes - cBytesReturned)

        End If

    Loop
End Function

' dirty method but effective ;)
Public Sub ResetACM()
    If hStream = INVALID_STREAM_HANDLE Then
        Exit Sub
    End If

    CloseConverter

    acmStreamOpen hStream, 0, _
                  ByVal clsStream.WFXPointer, wfxout, _
                  0, 0, 0, 0

    lngOutputLen = OUTPUT_BUFFER_MS / 1000 * wfxout.nAvgBytesPerSec
    acmStreamSize hStream, lngOutputLen, lngInputLen, ACM_STREAMSIZEF_DESTINATION

    ReDim intOutput(lngOutputLen / 2 - 1) As Integer
    ReDim btInput(lngInputLen - 1) As Byte

    blnEndOfStream = False
    blnFirst = True

    lngKeepInBuffer = 0
End Sub

Public Function CloseConverter() As Boolean
    On Error Resume Next

    acmStreamClose hStream, 0
    hStream = INVALID_STREAM_HANDLE

    ZeroMemory btInput(0), lngInputLen
    ZeroMemory intOutput(0), lngOutputLen

    blnEndOfStream = False
    lngBufferData = 0
    lngPosInBuffer = 0
    lngInputLen = 0
    lngOutputLen = 0
    lngKeepInBuffer = 0

    CloseConverter = True
End Function

Public Function InitConversion( _
    ByRef stream As ISoundStream, _
    ByRef DSPCallback As IDSPCallback, _
    Optional Force16Bit As Boolean = False _
) As Boolean

    Dim mmr     As Long

    ' create input/output WFX
    DXCopyMemory wfxinp, ByVal stream.WFXPointer, Len(wfxinp)

    If wfxinp.wBitsPerSample < 8 Then
        wfxinp.wBitsPerSample = 8
    ElseIf wfxinp.wBitsPerSample > 8 Then
        wfxinp.wBitsPerSample = 16
    End If

    If Force16Bit Or stream.Info.Force16Bit Then
        wfxinp.wBitsPerSample = 16
    End If

    With wfxinp
        wfxout = CreateWFX(.nSamplesPerSec, .nChannels, .wBitsPerSample)
    End With

    ' create the ACM stream
    mmr = acmStreamOpen(hStream, 0, _
                        ByVal stream.WFXPointer, wfxout, _
                        0, 0, 0, 0)

    If mmr <> 0 Then
        If Force16Bit Then Exit Function

        ' could not open the ACM stream,
        ' try with an other bps
        If wfxout.wBitsPerSample = 16 Then
            wfxout.wBitsPerSample = 8
        Else
            wfxout.wBitsPerSample = 16
        End If

        mmr = acmStreamOpen(hStream, 0, _
                            ByVal stream.WFXPointer, wfxout, _
                            0, 0, 0, 0)

        If mmr <> 0 Then Exit Function
    End If

    ' needed input size for OUT_BUFFER_MS
    ' milliseconds of PCM data
    lngOutputLen = OUTPUT_BUFFER_MS / 1000 * wfxout.nAvgBytesPerSec
    mmr = acmStreamSize(hStream, lngOutputLen, lngInputLen, ACM_STREAMSIZEF_DESTINATION)
    If mmr <> 0 Then
        acmStreamClose hStream, 0
        hStream = 0
        Exit Function
    End If

    ReDim intOutput(lngOutputLen / 2 - 1) As Integer
    ReDim btInput(lngInputLen - 1) As Byte

    Set clsStream = stream
    Set clsCB = DSPCallback

    blnEndOfStream = False
    blnFirst = True

    lngKeepInBuffer = 0

    InitConversion = True
End Function
