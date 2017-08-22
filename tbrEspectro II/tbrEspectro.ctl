VERSION 5.00
Begin VB.UserControl tbrEspectro 
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1875
   ScaleHeight     =   825
   ScaleWidth      =   1875
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "tbrEspectro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Dim CGMA As Single

Dim qMxL As Long
Dim qDisp As Long
Dim qModo As Long
Dim qFX_Enabled As Boolean

Dim qSensibilidad As Double

Dim qBCol As Long
Dim qCol As Long

Dim qDoCls As Boolean

Dim qBlur As Boolean
Dim qLuz As Boolean
Dim qLineas As Boolean

Dim qRefreshRate As Long

Private WithEvents clsRecorder  As WaveInRecorder
Attribute clsRecorder.VB_VarHelpID = -1
Private intSamples()            As Integer
Private clsVis                  As clsDraw

Public Property Get RefreshRate() As Long
    RefreshRate = qRefreshRate
End Property

Public Property Let RefreshRate(ByVal vNewValue As Long)
    qRefreshRate = vNewValue
    Timer1.Interval = qRefreshRate
End Property

Public Property Get Blur() As Boolean
    Blur = qBlur
End Property

Public Property Let Blur(ByVal vNewValue As Boolean)
    qBlur = vNewValue
End Property

Public Property Get Luz() As Boolean
    Luz = qLuz
End Property

Public Property Let Luz(ByVal vNewValue As Boolean)
    qLuz = vNewValue
End Property

Public Property Get Lineas() As Boolean
    Lineas = qLineas
End Property

Public Property Let Lineas(ByVal vNewValue As Boolean)
    qLineas = vNewValue
End Property

Public Property Get DoCls() As Boolean
    DoCls = qDoCls
End Property

Public Property Let DoCls(ByVal vNewValue As Boolean)
    qDoCls = vNewValue
End Property

Public Property Get Sensibilidad() As Double
    Sensibilidad = qSensibilidad
End Property

Public Property Let Sensibilidad(vSens As Double)
    qSensibilidad = vSens
End Property

Public Property Get ColorLinea() As OLE_COLOR
    ColorLinea = qCol
End Property

Public Property Let ColorLinea(bColor As OLE_COLOR)
    qCol = bColor
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = qBCol
End Property

Public Property Let BackColor(bColor As OLE_COLOR)
    qBCol = bColor
    P.BackColor = bColor
End Property

Private Sub clsRecorder_GotData(intBuffer() As Integer, lngLen As Long)
    intSamples = intBuffer
End Sub

Public Sub Comenzar()
    If Not clsRecorder.StartRecord("44100", 2) Then
        MsgBox "No se puede comunicar con dispositivo", vbExclamation, "tbrEspectro"
    End If
    Timer1.Enabled = True
End Sub

Public Sub Detener()
    If Not clsRecorder.StopRecord Then
        MsgBox "No se puede detener", vbExclamation
    End If
    Timer1.Enabled = False
End Sub

Private Sub DispositivoX(CualDisp As Long)
    If Not clsRecorder.SelectDevice(CualDisp) Then
        MsgBox "Ocurrio un error con este dispositivo!", vbExclamation, "tbrEspectro"
        Exit Sub
    End If
    'CargoLineas
End Sub
Private Sub MixerLinex(CualMxL As Long)
    If Not clsRecorder.SelectMixerLine(CualMxL) Then
        MsgBox "No se puede seleccionar este dispositivo!", vbExclamation, "Telefono"
    End If
    'CargoLineas
End Sub

Private Sub Timer1_Timer()
    If qModo = 1 Then
        clsVis.DrawFrequencies intSamples, P, qCol, qSensibilidad, qDoCls
    Else
        clsVis.DrawAmplitudes intSamples, P, qCol, qBCol
    End If
    
    'CGMA = CGMA + 10
    'If CGMA > 40 Then
    '    CGMA = -40
    'End If
    
    'fxBlur P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15)
    'fxEngrave P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15), 5
    'fxMosaic P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15), 2
    'fxGridelines P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15), vbBlack, 200, 1
    'fxRelief P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15)
    
    If qFX_Enabled = True Then
        If qBlur = True Then
            fxBlur P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15)
        End If
        If qDoCls = True Then
            If qLuz = True Then
                fxLight P.hdc, 10, 10, vbWhite, 100, 50, 150
            End If
        End If
        If qLineas = True Then
            fxScanlines P.hdc, 0, 0, P.Width / 15, P.Height / 15, P.hdc, 0, 0, (P.Width / 15), (P.Height / 15), vbBlack, 200, 1, True, False
        End If
    End If
    '¯
    
    'fxScanlines
    'P2.PaintPicture P.Image, 0, 0, P2.Width / 15, P2.Height / 15
End Sub

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

Private Sub UserControl_Initialize()
    'CGMA = -250
    Set clsRecorder = New WaveInRecorder
    Set clsVis = New clsDraw
    'CargoDispositivos
    ReDim intSamples(FFT_SAMPLES - 1) As Integer
    qModo = 1
    qFX_Enabled = True
    DispositivoX 0
    qBCol = &H800080
    qCol = vbYellow
    qSensibilidad = 1
    qDoCls = True
    qBlur = True
    qLuz = True
    qLineas = True
    qMxL = 0
    qDisp = 0
    qRefreshRate = 25
    'Comenzar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    qDisp = PropBag.ReadProperty("xDisp"): DispositivoX qDisp
    qModo = PropBag.ReadProperty("xModo"): SetModo qModo
    qFX_Enabled = PropBag.ReadProperty("xFX")
    qMxL = PropBag.ReadProperty("xMxL")
    
    qBCol = PropBag.ReadProperty("xBCol"): P.BackColor = qBCol
    qCol = PropBag.ReadProperty("xCol")
    
    qSensibilidad = PropBag.ReadProperty("xSense")
    qDoCls = PropBag.ReadProperty("xDoCls")
    
    qBlur = PropBag.ReadProperty("xBlur")
    qLuz = PropBag.ReadProperty("xLuz")
    qLineas = PropBag.ReadProperty("xLineas")
    
    qRefreshRate = PropBag.ReadProperty("xRefreshRate")

    'LS2.BackColor = PropBag.ReadProperty("fBColor")

End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 200 Then Exit Sub
    If UserControl.Height < 200 Then Exit Sub
    P.Width = UserControl.Width
    P.Height = UserControl.Height
    DRW_BARWIDTH = (UserControl.Width / 15) / 20
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "xDisp", qDisp
    PropBag.WriteProperty "xModo", qModo
    PropBag.WriteProperty "xFX", qFX_Enabled
    PropBag.WriteProperty "xMxL", qMxL
    
    PropBag.WriteProperty "xBCol", qBCol
    PropBag.WriteProperty "xCol", qCol
    
    PropBag.WriteProperty "xSense", qSensibilidad
    PropBag.WriteProperty "xDoCls", qDoCls

    PropBag.WriteProperty "xBlur", qBlur
    PropBag.WriteProperty "xLuz", qLuz
    PropBag.WriteProperty "xLineas", qLineas
    
    PropBag.WriteProperty "xRefreshRate", qRefreshRate

    'PropBag.WriteProperty "fBColor", LS2.BackColor
End Sub

Public Property Get DispositivoNumero() As Long
    DispositivoNumero = qDisp
End Property

Public Property Let DispositivoNumero(ByVal vNewValue As Long)
    qDisp = vNewValue
    DispositivoX vNewValue
End Property

Public Property Get Modo() As Long
    Modo = qModo
End Property

Public Property Let Modo(ByVal vNewValue As Long)
    qModo = vNewValue
    SetModo vNewValue
End Property

Public Property Get FX_Enabled() As Boolean
    FX_Enabled = qFX_Enabled
End Property

Public Property Let FX_Enabled(ByVal vNewValue As Boolean)
    qFX_Enabled = vNewValue
End Property

Private Sub SetModo(X As Long)
    P.Cls
    Select Case X
        Case 1
            P.AutoRedraw = True
        Case 2
            P.AutoRedraw = False
    End Select
End Sub

Public Property Get Dispositivo() As Long
    Dispositivo = qDisp
End Property

Public Property Let Dispositivo(ByVal vNewValue As Long)
    qDisp = vNewValue
    DispositivoX vNewValue
End Property

Public Property Get MIXERLINE() As Long
    MIXERLINE = qMxL
End Property

Public Property Let MIXERLINE(ByVal vNewValue As Long)
    qMxL = vNewValue
    MixerLinex vNewValue
End Property

