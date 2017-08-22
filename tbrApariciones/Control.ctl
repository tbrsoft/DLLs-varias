VERSION 5.00
Begin VB.UserControl tbrApariciones 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Control.ctx":0000
   ScaleHeight     =   540
   ScaleWidth      =   540
   ToolboxBitmap   =   "Control.ctx":0C42
   Begin VB.Timer T 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   300
      Top             =   150
   End
End
Attribute VB_Name = "tbrApariciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event Temino(QModo As Long)

Dim qHeight As Long
Dim qWidth As Long
Dim qX As Long
Dim qY As Long
Dim qPicIn As Object
Dim qPicOut As Object

Dim Modo As Boolean
Dim C As Long

Dim qColorTransparente As OLE_COLOR

Dim qTimeInterval As Long
Dim qFrameRate As Long

Private Sub T_Timer()
    If Modo Then
        C = C + qFrameRate
    Else
        C = C - qFrameRate
    End If
    qPicOut.Cls
    
    If qHeight = -1 Then
        fxAlphaBlend qPicOut.hDC, qX, qY, qPicIn.Width / 15, qPicIn.Height / 15, qPicIn.hDC, 0, 0, qPicIn.Width / 15, qPicIn.Height / 15, C, qColorTransparente
    Else
        fxAlphaBlend qPicOut.hDC, qX, qY, qWidth, qHeight, qPicIn.hDC, 0, 0, qPicIn.Width / 15, qPicIn.Height / 15, C, qColorTransparente
    End If
    
    qPicOut.Refresh
    If C > 255 Then
        RaiseEvent Temino(CLng(Modo))
        T.Enabled = False
    End If
    If C < 0 Then
        RaiseEvent Temino(CLng(Modo))
        T.Enabled = False
    End If
End Sub

Private Sub UserControl_Initialize()
    qTimeInterval = 100
    qFrameRate = 15
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Local Error Resume Next
    qTimeInterval = PropBag.ReadProperty("xTime")
    qFrameRate = PropBag.ReadProperty("xFrame")
    qColorTransparente = PropBag.ReadProperty("xColoT")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "xTime", qTimeInterval
    PropBag.WriteProperty "xFrame", qFrameRate
    PropBag.WriteProperty "xColoT", qColorTransparente
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width <> 540 Then UserControl.Width = 540
    If UserControl.Height <> 540 Then UserControl.Height = 540
End Sub

Public Property Get TimeInterval() As Long
    TimeInterval = qTimeInterval
End Property

Public Property Let TimeInterval(ByVal vNewValue As Long)
    qTimeInterval = vNewValue
End Property

Public Property Get FrameRate() As Long
    FrameRate = qFrameRate
End Property

Public Property Let FrameRate(ByVal vNewValue As Long)
    qFrameRate = vNewValue
End Property

Public Property Get ColorTransparente() As OLE_COLOR
    ColorTransparente = qColorTransparente
End Property

Public Property Let ColorTransparente(ByVal vNewValue As OLE_COLOR)
    qColorTransparente = vNewValue
End Property

'Si no se aclara Width o Height entonces se usa el alto y ancho de la ImagenFuente
Public Sub Aparecer(ImagenDestino As Object, ImagenFuente As Object, X As Long, Y As Long, CModo As Boolean, Optional HeightPixel As Long = -1, Optional WidthPixel As Long = -1, Optional Auxiliar As String = "")
    Modo = CModo
    If Modo Then
        C = 0
    Else
        C = 255
    End If
    
    qHeight = HeightPixel
    qWidth = WidthPixel
    qX = X
    qY = Y
    Set qPicOut = ImagenDestino
    Set qPicIn = ImagenFuente
    T.Interval = qTimeInterval
    T.Enabled = True
End Sub

Public Sub MarcaDeAgua(ImagenDestino As Object, ElColor As Long, Level As Long)
    Set qPicOut = ImagenDestino
    'fxAmbientLight qPicOut.hDC, 0, 0, qPicOut.Width / 15, qPicOut.Height / 15, qPicOut.hDC, 0, 0, qPicOut.Width / 15, qPicOut.Height / 15, ElColor, Level
    fxTone qPicOut.hDC, 0, 0, qPicOut.Width / 15, qPicOut.Height / 15, qPicOut.hDC, 0, 0, qPicOut.Width / 15, qPicOut.Height / 15, ElColor, Level
    qPicOut.Refresh
    Set qPicOut.Picture = qPicOut.Image
End Sub

