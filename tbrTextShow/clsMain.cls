VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private TextEx As New cTextEx
Private StdFontEx As New cStdFontEx

Public Sub Iniciar(FontName As String, FontSize As Long, FontBold As Boolean, FontItalic As Boolean, FontStrikethrough As Boolean, FontUnderLine As Boolean, FontColour As Long)
  With StdFontEx
    .Name = FontName
    .Size = FontSize
    .Bold = FontBold
    .Italic = FontItalic
    .Strikethrough = FontStrikethrough
    .UnderLine = FontUnderLine
    .Charset = 0
    .Colour = FontColour
  End With
End Sub
Public Sub Dibujar(Texto As String, qHDC As Long, X As Long, Y As Long, Ancho As Long, Alto As Long)

  Dim mRECT As RECT
  
  
  With TextEx
    .RectLeft = X
    .RectRight = Ancho
    .RectTop = Y
    .RectBottom = Alto
    .Draw StdFontEx, , , , , qHDC, Texto, 0
  End With
  
End Sub

