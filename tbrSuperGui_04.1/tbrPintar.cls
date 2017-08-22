VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrPintar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Const STRETCHMODE = vbPaletteModeNone

Public Sub tbrPintarX(PicIn As Object, PicOut As Object, X As Long, Y As Long, Ancho As Long, Alto As Long, Optional Pixel As Boolean = False)
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
Public Sub tbrPintarInHDC(HdcIn As Long, HdcOut As Long, X As Long, Y As Long, AnchoOut As Long, AltoOut As Long, AnchoIn As Long, AltoIn As Long)
    Call SetStretchBltMode(HdcOut, STRETCHMODE)
    Call StretchBlt(HdcOut, X, Y, AnchoOut, AltoOut, HdcIn, 0, 0, AnchoIn, AltoIn, vbSrcCopy)
End Sub

