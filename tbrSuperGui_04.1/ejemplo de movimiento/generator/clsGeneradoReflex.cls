VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeneradoReflex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim DestObjHdc As Long

Public Event CargandoArchivo(qArchivo As String, qPorcentajeVoy As Long)

Private Type RectaHDC
    rHDC As Long
    rX As Long
    rY As Long
    rAncho As Long
    rAlto As Long
End Type

Dim ReflejoIn As RectaHDC
Dim ReflejoOut As RectaHDC

Const AC_SRC_OVER = &H0

Private Type AlphaOptions
  AlphaOption As Byte
  AlphaFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long

'====================================================================
'Dim SpriteHDC As Long
Dim FondoHDC As Long

Dim hdcFondo As New tbrAlphaB

Dim qAnchoTotal As Long
Dim qAltoTotal As Long

Dim qCantidadImagenes As Long

Dim qFondoColor As Long

Dim qPorcentajeReflejo As Long
Dim qMargenPixeles As Long
Dim qAnchoCadaImagen As Long
Dim qAltoCadaImagen As Long

Private Sub IniHDC(CualHDC As Long, qAncho As Long, qAlto As Long)
    Dim TempBMP As Long
    Dim ObjCreado2 As Long

    TempBMP = CreateCompatibleBitmap(DestObjHdc, qAncho, qAlto)
    CualHDC = CreateCompatibleDC(0)
    ObjCreado2 = SelectObject(CualHDC, TempBMP)

    DeleteObject TempBMP
    DeleteObject ObjCreado2
End Sub

Public Sub GenerarImagen(lstPathImagenes() As String, AnchoCadaImagen As Long, AltoCadaImagen As Long, qAnyPictureHDC As Long)
    Dim i As Long
    Dim tmpHDC As Long
    Dim fX As Long
    Dim fY As Long
    Dim fAncho As Long
    Dim fAlto As Long
    
    DestObjHdc = qAnyPictureHDC
    
    qAnchoCadaImagen = AnchoCadaImagen
    qAltoCadaImagen = AltoCadaImagen
    qCantidadImagenes = UBound(lstPathImagenes)
    
    qFondoColor = RGB(200, 10, 10)
    qMargenPixeles = 5
    qPorcentajeReflejo = 60
    
    qAnchoTotal = ((qAnchoCadaImagen + qMargenPixeles) * qCantidadImagenes) + (qMargenPixeles)
    qAltoTotal = qAltoCadaImagen + (qMargenPixeles * 2) + (qAltoCadaImagen * (qPorcentajeReflejo / 100))

    'Crear el Fondo donde se van a poner la imagenes
    IniHDC FondoHDC, qAnchoTotal, qAltoTotal
    'PintarHDCColor FondoHDC, qFondoColor, qAnchoTotal, qAltoTotal
    
    'Creo el HDC donde se van cargando las imagenes
    'IniHDC SpriteHDC, qAnchoCadaImagen, qAltoCadaImagen
    
    'Pinto todas las imagenes en el FondoHDC
    fX = qMargenPixeles
    fY = qMargenPixeles
    For i = 1 To qCantidadImagenes
        CargarImagen lstPathImagenes(i), tmpHDC, fAncho, fAlto
        
        tbrPintarInHDC tmpHDC, FondoHDC, fX, fY, qAnchoCadaImagen, qAltoCadaImagen, fAncho, fAlto
        
        '===========================
        'TEST/Comienzo
        '===========================
        'Ok
        'tbrPintarInHDC tmpHDC, Form1.picTMP.hdc, 0, 0, qAnchoCadaImagen, qAltoCadaImagen, fAncho, fAlto
        
        'Ok
        'tbrPintarInHDC FondoHDC, Form1.picTMP.hdc, 0, 0, 500, 300, qAnchoTotal, qAltoTotal
        
        'Form1.picTMP.Refresh
        '===========================
        'TEST/Fin
        '===========================
        fX = fX + qAnchoCadaImagen + qMargenPixeles
        ret = DeleteDC(tmpHDC)
        
        RaiseEvent CargandoArchivo(lstPathImagenes(i), (i * 100) / qCantidadImagenes)
        DoEvents
    Next i

    'Creo el reflejo
    CrearReflejo
End Sub

Private Sub PintarHDCColor(fHDC As Long, fColor As Long, fAncho As Long, fAlto As Long)
    hdcFondo.CerrarGraficos
    hdcFondo.CrearHDC_Color fColor
    hdcFondo.MostrarAlpha fHDC, 0, 0, 100, fAncho, fAlto
End Sub

Private Sub CrearReflejo()
    Dim PorcentajeAlpha As Single
    Dim fY_comienza As Long
    Dim fY_termina As Long
    Dim i As Long
    
    'fY_comienza = La parte inferior de las imagenes (sobra el espacio reservado para el refejo)
    'fY_termina = El final de todo el espacio reservado para el refeljo
    
    fY_comienza = qAltoCadaImagen + qMargenPixeles
    fY_termina = fY_comienza + (qAltoCadaImagen * (qPorcentajeReflejo / 100))
    
    'Aqui hago un 'for' desde arriba para abjo de toda la imagen en general
    'y la voy pintando cada vez mas transparente. De esa manera creo un reflejo.
    ReflejoIn.rHDC = FondoHDC
    ReflejoOut.rHDC = FondoHDC
    For i = fY_comienza To fY_termina
        'Calculo Imagen que se copia.
        ReflejoIn.rX = 0
        ReflejoIn.rY = (fY_comienza) - (i - fY_comienza)
        ReflejoIn.rAncho = qAnchoTotal
        ReflejoIn.rAlto = 1
        'Calculo Imagen de reflejo.
        ReflejoOut.rX = 0
        ReflejoOut.rY = i + qMargenPixeles
        ReflejoOut.rAncho = qAnchoTotal
        ReflejoOut.rAlto = 1
        
        'Calculo la cantidad de semitransparencia
        PorcentajeAlpha = (fY_termina - i) / (fY_termina - fY_comienza)
        PorcentajeAlpha = 130 * PorcentajeAlpha
        'Pinto el pixel de alto con la semitransparencia
        MostrarAlpha ReflejoIn, ReflejoOut, CLng(PorcentajeAlpha)
    Next i
End Sub

Private Sub MostrarAlpha(RectaIn As RectaHDC, RectaOut As RectaHDC, qAlpha254 As Long)
    Dim AO As AlphaOptions
    Dim ptAO As Long
    
    If qAncho = -1 Then
        qAncho = Ancho
    End If
    If qAlto = -1 Then
        qAlto = Alto
    End If
    
    With AO
        .AlphaOption = AC_SRC_OVER
        .AlphaFlags = 0
        .SourceConstantAlpha = qAlpha254
        .AlphaFormat = 0
    End With
    
    RtlMoveMemory ptAO, AO, 4
    AlphaBlend RectaOut.rHDC, RectaOut.rX, RectaOut.rY, RectaOut.rAncho, RectaOut.rAlto _
                        , RectaIn.rHDC, RectaIn.rX, RectaIn.rY, RectaIn.rAncho, RectaIn.rAlto, ptAO
End Sub

Public Sub GrabarEnArchivo(Archivo As String, tmpPic As PictureBox)
    tmpPic.Width = qAnchoTotal
    tmpPic.Height = qAltoTotal

    BitBlt tmpPic.hdc, 0, 0, qAnchoTotal, qAltoTotal, FondoHDC, 0, 0, SRCCOPY
    tmpPic.Refresh
    
    SavePicture tmpPic.Image, Archivo
End Sub


Public Sub CerrarGraficos()
    Dim ret As Long
    ret = DeleteDC(FondoHDC)
End Sub

