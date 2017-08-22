VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrMP3EncDll"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Public Event Estado(Porcentaje As Integer)

Public Enum BitRates
    a_Kbps_320 = 320
    b_Kbps_256 = 256
    c_Kbps_224 = 224
    d_Kbps_192 = 192
    e_Kbps_160 = 160
    f_Kbps_128_Default = 128
    g_Kbps_112 = 112
    h_Kbps_96 = 96
    i_Kbps_80 = 80
    j_Kbps_64 = 64
    k_Kbps_56 = 56
    l_Kbps_48 = 48
    m_Kbps_40 = 40
End Enum

'API declarations for encoding wrapper
Private Declare Function SetBitrate Lib "MEncoder.dll" (ByVal bit As Integer) As Long
Private Declare Function EncodeMp3 Lib "MEncoder.dll" (ByVal lpszWavFile As String, lpCallback As Any) As Long
Private Declare Function SetCopyright Lib "MEncoder.dll" (ByVal cpy As Boolean) As Long
Private Declare Function SetOriginal Lib "MEncoder.dll" (ByVal org As Boolean) As Long
Private Declare Function SetCRC Lib "MEncoder.dll" (ByVal crc As Boolean) As Long
Private Declare Function SetPrivate Lib "MEncoder.dll" (ByVal priv As Boolean) As Long
Private Declare Function Cancel Lib "MEncoder.dll" (ByVal cncl As Boolean) As Long

Private Sub Class_Initialize()

    tErr.FileLog = App.Path + "\logMp3Enc.log"

    Call SetBitrate(128)
    
    tErr.Anotar "ini-mp3Enc-11"
    Call SetCopyright(False)
    
    tErr.Anotar "ini-mp3Enc-12"
    Call SetOriginal(False)
    
    tErr.Anotar "ini-mp3Enc-13"
    Call SetCRC(False)
    
    tErr.Anotar "ini-mp3Enc-14"
    Call SetPrivate(False)
    
    tErr.Anotar "ini-mp3Enc-15"
End Sub

Public Sub SetLog(sPath As String)
    tErr.FileLog = sPath
End Sub

Public Sub Iniciar(BitRate As BitRates, Optional Set_Copyright As Boolean = False, _
    Optional Set_Original As Boolean = False, Optional Set_CRC As Boolean = False, _
    Optional Set_Private As Boolean = False)
    
        On Local Error GoTo ERRINI
    
        tErr.Anotar "ini-mp3Enc-1"
    
        Call SetBitrate(BitRate)
        
        tErr.Anotar "ini-mp3Enc-2"
        Call SetCopyright(Set_Copyright)
        
        tErr.Anotar "ini-mp3Enc-3"
        Call SetOriginal(Set_Original)
        
        tErr.Anotar "ini-mp3Enc-4"
        Call SetCRC(Set_CRC)
        
        tErr.Anotar "ini-mp3Enc-5"
        Call SetPrivate(Set_Private)
        
        tErr.Anotar "ini-mp3Enc-6"
        Exit Sub
        
ERRINI:
    tErr.AppendLog "ErrIni-mp3Enc-091", tErr.ErrToTXT(Err)
    Resume Next
End Sub

Public Function Encode(FileIn As String) As Integer
    Dim nRes As Integer
    'Now start emcpding the file
    nRes = EncodeMp3(FileIn, AddressOf EnumEncoding) 'supongo que le manda los eventos a enumEncoding de status
    
    Encode = nRes
    
    'User Notifications
    '<> -1 "MP3 encoding complete"
    '   -2 "Encoding stopped by user"
    'Else  "Encoding failed", vbExclamation, App.Title
    
    If nRes = -1 Then
        tErr.AppendLog "Enconde-Error"
    End If
End Function

Public Sub CancelEncode()
    tErr.Anotar "ini-mp3Enc-8"
    Call Cancel(True)
End Sub

'Public Function CambiaEstado() As Integer
'    RaiseEvent Estado(PcjeActual)
'End Function

Public Function GetEventos() As clsLlamarEvento
    On Local Error GoTo ErrEve
    
    tErr.Anotar "ini-mp3Enc-9"
    Set GetEventos = Llamador
    
    tErr.Anotar "ini-mp3Enc-10"
    
    Exit Function
    
ErrEve:
    tErr.AppendLog "errEve", tErr.ErrToTXT(Err)
End Function
