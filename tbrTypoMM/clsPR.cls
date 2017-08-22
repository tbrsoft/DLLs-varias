VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'los tipos tienen que manegar cada uno sus precios
'aqui se maneja todo eso

'cada contenido puede tener uno o mas precios
'por ahora tengo
'1- precio por reproduccion (se refiere a definiciones generales de todas las promociones)

'cuando se busca precios de fonolas si o si se pide de a uno y es posible que sea parte de una promocion
'por ejemplo puso $ 2 y eso te daba 3 canciones. El tipo no pide las tres canciones juntas
'pide de a una, para eso hay un "precio actual" que va llevando adelante el extraño calculo necesario)

'2- Precio actual, luego de calculos muy extraños es el que realmente hay que cobrar ahora

'3- precio por descarga
'4- precio vip (para que se adelante a la reproduccion actual)
'5- precio por preview (en cero me imagino pero simepre hay algun desubicado)
'activarla implicaria activar dos sistemas de precios como en la reproduccion, uno que sea base y otro que sea
'el precio actual del preview


Public PricePlay As New modPrecios
Public PriceActual As New modPrecios
Public PriceVIP As New modPrecios
Public PriceDescarga As New modPrecios

'Public PricePreviewPlay As tbrPrecioRK.modPrecios
'Public PricePreviewActual As tbrPrecioRK.modPrecios


'actualizar el precioActual en base a la lista de precios de reproduccion
'esto se hace cada vez que varian los creditos del sistema
Public Function UpdatePriceBase(VC As Single, Cred As Single) As Long
    UpdatePriceBase = PriceActual.UpdatePrice(PricePlay, VC, Cred)
End Function

Private Sub Class_Initialize()
    'para despues poder leerlo segun que tipo son
    PricePlay.TipoPrecio = Tipo_PrecioPlay
    PriceActual.TipoPrecio = Tipo_PreciosActual
    PriceVIP.TipoPrecio = Tipo_PrecioVIP
    PriceDescarga.TipoPrecio = Tipo_PreciosDescarga
End Sub
