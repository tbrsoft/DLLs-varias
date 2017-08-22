VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrWR2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim NuVumPromedio As Single

Private xDispositivo As Long
Private xLinea As Long
Private xArchivo As String
'-------------------------

Private Const FFT_SAMPLES            As Long = 1024

Private intSamples()            As Integer
Private lngBytesPerSec          As Long
Private WithEvents clsRecorder  As WaveInRecorder
Attribute clsRecorder.VB_VarHelpID = -1
Private clsEncoder              As EncoderWAV
Private clsDSP                  As clsDSP

Private terr As New tbrErrores.clsTbrERR
Private IsSaving As Long 'boolean es muy limitado


Private Sub Class_Initialize()
    
    terr.FileLog = App.path + "\logWII.txt"
    terr.LargoAcumula = 950
    
    Set clsDSP = New clsDSP
    terr.Anotar "baa"
    
    Set clsEncoder = New EncoderWAV
    terr.Anotar "bab"
    
    Set clsRecorder = New WaveInRecorder
    terr.Anotar "bac"
    
    
    
    xDispositivo = 0
    xLinea = 2
    xArchivo = Left(App.path, 1) + ":\tbrWaveRecord.wav"  'valor predeterminado
    terr.Anotar "bad"
    
    IsSaving = 0
End Sub

Public Sub SetFileLog(nF As String)
    terr.FileLog = nF
End Sub

Private Sub clsRecorder_GotData(intBuffer() As Integer, lngLen As Long)
    
    ' save the current buffer for visualizing it
    intSamples = intBuffer
    
    clsDSP.ProcessSamples intSamples
    
    ' GotData could also be raised after recording
    ' got stopped because a buffer was just finished
    ' when StopRecord got called
    If Not clsRecorder.IsRecording Then Exit Sub
    
    ' update recorded time
    lngMSEncoded = lngMSEncoded + ((lngLen / lngBytesPerSec) * 1000)
    
    If Not clsEncoder Is Nothing Then
        ' send PCM data to the WAV encoder
        If clsEncoder.Encoder_Encode(VarPtr(intSamples(0)), lngLen, 0) = SND_ERR_WRITE_ERROR Then
            'DETENER aqui
            MsgBox "Ocurrio un Error, revise que tiene espacio suficiente en el Disco", vbExclamation, "tbrWaveRecord"
        End If
    End If
    
    Dim Vums As Single
    Vums = 0
    For i = 0 To UBound(intSamples)
        Vums = Vums + (Abs(intSamples(i) / 32768#))
    Next i
    NuVumPromedio = Vums / UBound(intSamples)
    
    'NuVumPromedio = Abs(intSamples(0)) / 32768#
End Sub

Public Function Grabar(Optional qChannels As Integer = 2, Optional qDispositivo As Long = -1, Optional RecFile As Boolean = True) As Long
    'lngBytesPerSec = (CLng(cboSamplerate.Text) * (2 * (chkStereo.value + 1))) 'BRR
    lngBytesPerSec = (44100 * (2 * 2)) 'BRR
    
    clsDSP.samplerate = 44100
    clsDSP.Channels = qChannels
    
    e = clsEncoder.Encoder_EncoderInit(44100, qChannels, xArchivo)
    If Not clsRecorder.StartRecord(44100, qChannels, qDispositivo) Then
        Grabar = 1
        IsSaving = 0
        terr.AppendLog "No se puede empezar a grabar"
    Else
        Grabar = 0
        IsSaving = 1
    End If
End Function

Public Function IsSavingWAV() As Long
    IsSavingWAV = IsSaving
End Function

Public Function Detener() As Long
    'clsencoder.
    clsEncoder.Encoder_EncoderClose
    If Not clsRecorder.StopRecord Then
        terr.AppendLog "No se puede detener grabación !"
        Detener = 1
    Else
        IsSaving = 0
        Detener = 0
    End If
    'lblArch = NombreNoRepe("C:\LLAMADAS\Cliente.wav")
End Function

Public Property Get Dispositivo() As Long
    Dispositivo = xDispositivo
End Property

Public Property Let Dispositivo(ByVal vNewValue As Long)
    terr.Anotar "aaa", vNewValue, clsRecorder.DeviceName(vNewValue)
    xDispositivo = vNewValue
    SetDispositivo vNewValue
End Property

Public Property Get Linea() As Long
    Linea = xLinea
End Property

Public Property Let Linea(ByVal vNewValue As Long)
    terr.Anotar "aab", vNewValue, clsRecorder.MixerLineName(vNewValue)
    
    If SetLinea(vNewValue) = 0 Then
        xLinea = vNewValue
    Else
        'LINEA NO VALIDA !?
        terr.AppendSinHist vbCrLf + "**Lnv:0102->"
        
        xLinea = -1 'MARCAR QUE FALLO!!!
    End If
End Property

Public Property Get Archivo() As String
    Archivo = xArchivo
End Property

Public Property Let Archivo(ByVal vNewValue As String)
    terr.Anotar "aac", vNewValue
    xArchivo = vNewValue
End Property

Public Sub LogDispositivos()
    terr.Anotar "aad00:00"
    Dim sp() As String
    For i = 0 To clsRecorder.DeviceCount - 1
        sp = Split(clsRecorder.DeviceFullData(i), Chr(5))
        For k = 0 To UBound(sp)
            terr.Anotar "aad55", i, k, sp(k)
        Next k
    Next
    'terr.AppendLog ":PLACAS-LogDisp:"
End Sub

Public Sub LogLineas()
    terr.Anotar "aae00:00"
    For i = 0 To clsRecorder.MixerLineCount - 1
        terr.Anotar "aae", i, clsRecorder.MixerLineName(i)
    Next
    'terr.AppendLog ":LINEAS:"
End Sub

Public Sub CargoDispositivos(lLista As Object)
    lLista.Clear
    For i = 0 To clsRecorder.DeviceCount - 1
        terr.Anotar "aad44", i, clsRecorder.DeviceName(i)
        lLista.AddItem clsRecorder.DeviceName(i)
    Next
    terr.AppendLog ":PLACAS-CargoDisp:"
End Sub

Public Sub CargoLineas(lLista As Object)
    lLista.Clear
    For i = 0 To clsRecorder.MixerLineCount - 1
        terr.Anotar "aae", i, clsRecorder.MixerLineName(i)
        lLista.AddItem clsRecorder.MixerLineName(i)
    Next
    terr.AppendLog ":LINEAS:"
End Sub

Private Sub SetDispositivo(CualDisp As Long)
    terr.Anotar "aaf", CualDisp, clsRecorder.DeviceName(CualDisp)
    If Not clsRecorder.SelectDevice(CualDisp) Then
        terr.AppendLog "Ocurrio un error con este dispositivo!"
        Exit Sub
    End If
    terr.AppendLog "DISP(" + CStr(CualDisp) + ") " + clsRecorder.DeviceName(CualDisp)
    'CargoLineas
End Sub

Private Function SetLinea(CualMxL As Long) As Long
    terr.Anotar "aag", CualMxL, clsRecorder.MixerLineName(CualMxL)
    If Not clsRecorder.SelectMixerLine(CualMxL) Then
        terr.AppendLog "No se puede seleccionar este dispositivo!"
        SetLinea = 1
        Exit Function
    End If
    SetLinea = 0
    'CargoLineas
End Function

Public Property Get VumeterNum() As Single
    VumeterNum = NuVumPromedio
End Property


