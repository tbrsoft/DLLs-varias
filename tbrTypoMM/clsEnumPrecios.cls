VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnumPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***********************************
'**ENUMERACIONES********************
'***********************************

'tipos enumerados para ponerles un id y guiarme comodamento por ellos dentro y fuera de esta DLL
Public Enum MMTypes
    Error_MM = -1 'no lo uso todavia pero me va a servir seguramente
    NADA = 0 'para listatypos(0) u otras cosas extrañas
    Musica = 1
    Video = 2
    Karaoke = 3
    Ringtone = 4
    Wallpapers = 5
    Juegos_Java = 6
    ISO_NRG_CD = 7
    ISO_NRG_DVD = 8
    Videos_3GP = 9
    Temas_Movil = 10
    
    'Karaoke_Grabado
    'SEGUIRAQUI y hacer otro tipo (por cuestiones de extencion y tamañao me conviene que se califique como musica y salga a ese mismo precio)
End Enum

Public Type Extes
    sName As String
    isDupli As Boolean 'saber si otro tipo tambien la usa para tener en cuenta en la validaciones
End Type

Public Enum TipoPrecios
    Tipo_PrecioPlay = 1
    Tipo_PreciosDescarga = 2
    Tipo_PrecioVIP = 3
    Tipo_PreciosActual = 4 'es un calculo, derivado de precioPlay, no creo que lo use aqui
End Enum



'***********************************
'**FIN ENUMERACIONES****************
'***********************************
