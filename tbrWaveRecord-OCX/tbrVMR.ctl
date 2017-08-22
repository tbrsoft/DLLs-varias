VERSION 5.00
Begin VB.UserControl tbrWR 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "tbrVMR.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "tbrVMR.ctx":030A
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   465
      Left            =   15
      Top             =   15
      Width           =   465
   End
End
Attribute VB_Name = "tbrWR"
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

Private Sub UserControl_Initialize()
    Set clsDSP = New clsDSP
    Set clsEncoder = New EncoderWAV
    Set clsRecorder = New WaveInRecorder
    xDispositivo = 0
    xLinea = 2
    xArchivo = "C:\tbrWaveRecord.wav"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    xDispositivo = PropBag.ReadProperty("xDisp")
    xLinea = PropBag.ReadProperty("xLine")
    xArchivo = PropBag.ReadProperty("xArchiv")
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 480
    UserControl.Height = 480
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

Public Sub Grabar(Optional qChannels As Integer = 2, Optional qDispositivo As Long = -1, Optional RecFile As Boolean = True)
        'lngBytesPerSec = (CLng(cboSamplerate.Text) * (2 * (chkStereo.value + 1))) 'BRR
        lngBytesPerSec = (44100 * (2 * 2)) 'BRR
        
        clsDSP.samplerate = "44100"
        clsDSP.Channels = qChannels
        
        e = clsEncoder.Encoder_EncoderInit("44100", qChannels, xArchivo)
        If Not clsRecorder.StartRecord("44100", qChannels, qDispositivo) Then
            MsgBox "No se puede grabar!", vbExclamation, "tbrWaveRecord.ocx"
        End If
End Sub

Public Sub Detener()
    'clsencoder.
    clsEncoder.Encoder_EncoderClose
    If Not clsRecorder.StopRecord Then
        MsgBox "Could not stop recording!", vbExclamation
    End If
    'lblArch = NombreNoRepe("C:\LLAMADAS\Cliente.wav")
End Sub

Public Property Get Dispositivo() As Long
    Dispositivo = xDispositivo
End Property

Public Property Let Dispositivo(ByVal vNewValue As Long)
    xDispositivo = vNewValue
    SetDispositivo vNewValue
End Property

Public Property Get Linea() As Long
    Linea = xLinea
End Property

Public Property Let Linea(ByVal vNewValue As Long)
    xLinea = vNewValue
    SetLinea vNewValue
End Property

Public Property Get Archivo() As String
    Archivo = xArchivo
End Property

Public Property Let Archivo(ByVal vNewValue As String)
    xArchivo = vNewValue
End Property

Public Sub CargoDispositivos(lLista As Object)
    lLista.Clear
    For i = 0 To clsRecorder.DeviceCount - 1
        lLista.AddItem clsRecorder.DeviceName(i)
    Next
End Sub

Public Sub CargoLineas(lLista As Object)
    lLista.Clear
    For i = 0 To clsRecorder.MixerLineCount - 1
        lLista.AddItem clsRecorder.MixerLineName(i)
    Next
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "xDisp", xDispositivo
    PropBag.WriteProperty "xLine", xLinea
    PropBag.WriteProperty "xArchiv", xArchivo
End Sub

Private Sub SetDispositivo(CualDisp As Long)
    If Not clsRecorder.SelectDevice(CualDisp) Then
        MsgBox "Ocurrio un error con este dispositivo!", vbExclamation, "tbrWaveRecord"
        Exit Sub
    End If
    'CargoLineas
End Sub

Private Sub SetLinea(CualMxL As Long)
    If Not clsRecorder.SelectMixerLine(CualMxL) Then
        MsgBox "No se puede seleccionar este dispositivo!", vbExclamation, "Telefono"
    End If
    'CargoLineas
End Sub

Public Property Get VumeterNum() As Single
    VumeterNum = NuVumPromedio
End Property

