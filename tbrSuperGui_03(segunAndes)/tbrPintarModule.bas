Attribute VB_Name = "tbrPinta"
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Const STRETCHMODE = vbPaletteModeNone

Sub tbrPintar(PicIn As Object, PicOut As Object, X As Long, Y As Long, Ancho As Long, Alto As Long, Optional Pixel As Boolean = False)
    'Call SetStretchBltMode(picRSetMode.hdc, STRETCHMODE)
    'Call StretchBlt(picRSetMode.hdc, 0, 0, 70, 70, picOriginal.hdc, 0, 0, 100, 100, vbSrcCopy)
    
    If Pixel = True Then
        Call SetStretchBltMode(PicOut.hdc, STRETCHMODE)
        Call StretchBlt(PicOut.hdc, X, Y, Ancho, Alto, PicIn.hdc, 0, 0, (PicIn.Width), (PicIn.Height), vbSrcCopy)
    Else
        Call SetStretchBltMode(PicOut.hdc, STRETCHMODE)
        Call StretchBlt(PicOut.hdc, X, Y, Ancho, Alto, PicIn.hdc, 0, 0, (PicIn.Width / 15), (PicIn.Height / 15), vbSrcCopy)
    End If
    PicOut.Refresh

End Sub
Sub tbrPintarInHDC(HdcIn As Long, HdcOut As Long, X As Long, Y As Long, AnchoOut As Long, AltoOut As Long, AnchoIn As Long, AltoIn As Long)
    Call SetStretchBltMode(HdcOut, STRETCHMODE)
    Call StretchBlt(HdcOut, X, Y, AnchoOut, AltoOut, HdcIn, 0, 0, AnchoIn, AltoIn, vbSrcCopy)
End Sub

Public Sub CargarImagenEnHDC(Archivo As String, qHDC As Long)
    Dim ret As Long
    Dim AnchoDestino As Long
    Dim AltoDestino As Long
    
    Dim stdObjet As Long
    
    Dim lresult As Long
    
    Dim tmpName As String
    
    If Dir(Archivo) = "" Then Exit Sub
    
    Set stdPicAux_CI = LoadPicture(Archivo)
    lresult = SelectObject(qHDC, stdPicAux_CI.Handle)
    
    AnchoSprite = qAlgunFormulario.ScaleX(stdPicAux_CI.Width, vbHimetric, vbPixels)
    AltoSprite = qAlgunFormulario.ScaleX(stdPicAux_CI.Height, vbHimetric, vbPixels)
    
    'ElDestinoHDC = DestinoHDC
    
    If AnchoSprite < 1 Then Ancho = 1
    If AnchoSprite < 1 Then Alto = 1
    
'DESCARGAR MEMORIA:--------------------------
    ret = DeleteObject(lresult)
    ret = DeleteObject(stdPicAux_CI.Handle)
    ret = DeleteObject(stdPicAux_CI)
'--------------------------------------------
End Sub
