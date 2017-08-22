VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim pngClass As New LoadPNG

'piccargar es donde se carga la imagen. Picturebox o frm
    'el que sigue es el mismo tipo de objeto pero indica cual _
        es la imagen que esta atras para copiar los pixeles en las transparencias
    'pathIMG es el archivo PNG
    
    'la mejor opcion parece que ambas imagen sean la misma ya que:
    'pngClass.SetToBkgrnd False, 0, 0
    'si se pone asi escribe en las coordenadas 0,0 de piccargar el dibujo pero si o si toma
    'para poner de fondo los mismas coordenadas para los pixeles de la imagen del fondo
    'o sea que la imagen de arriba debe estar en 0,0 y no sirve. Si necesito poner 5 imagenes
    'en distintos lugares me caga. En cambio si uso:
    'pngClass.SetToBkgrnd TRUE, x, y
    'la imagen se escribe en la de fondo (no se necesita la primera) en las coordenadas que especifico
    'la contra es que no tengo un objeto PIC para por ejemplo usar mouse pointer
    'alli habría que usar algo transparente para reemplazar!

Public Function CargarPNG(Pic, PicAtras, pathIMG As String) As Long
                          
    On Local Error GoTo noLoad
    
    Dim P As PictureBox
    Set P = Pic
    
    P.Picture = LoadPicture("")
    pngClass.PicBox = P
    pngClass.SetToBkgrnd False, P.Left / 15, P.Top / 15, 0, 0 'set to Background (True or false), x and y
    pngClass.BackgroundPicture = PicAtras
    pngClass.SetAlpha = True 'when Alpha then alpha
    pngClass.SetTrans = True 'when transparent Color then transparent Color

    ' Visualiza el Archivo
    pngClass.OpenPNG pathIMG
    CargarPNG = 0 'no hay error
    
    Exit Function
    
noLoad:
    CargarPNG = Err.Number
    Exit Function
End Function
