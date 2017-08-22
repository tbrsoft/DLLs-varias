VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTypo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


'que se agreguen desde afuera a gusto !
'Public Enum Typos
'    'de reproduccion
'    rMusicaPlay = 1
'    rVideoPlay
'    rKaraokePlay
'    'de contenido extendido
'    rMusicaBuy 'la musica, videos y kar cuando se compra puede tener otro precio
'    rVideoBuy
'    rKaraokeBuy
'    rRingtones
'    rWallpapers
'    rJuegosJava
'    rISO_NRG_CD
'    rVideos_3GP
'    rTemasMovil
'    rISO_NRG_DVD
'End Enum

Public Enum ModoCompra
    CompraPaqueteArmado = 1
    PagaYDespuesUsa = 2
    Todos = 3
End Enum

Private mNombreBase As String 'por ejemplo musica
Private mNombre As String 'por ejemplo musicaPlay o MusicaPaquete
Private mDescripcion As String
Private mModoQueSeCompra As ModoCompra
Private mPublico As Long
Private mPrecioActual As Single 'en creditos
Private mPrecioActualVIP As Single 'en creditos
Private mCantidadPedida As Long 'para el carrito
Private mHelpUse As String 'ayuda para usarlo, por ejemplo juego java = revise su tamaño de pantalla y configuracion CLDC

Public Property Get Nombre() As String
    Nombre = mNombre
End Property

Public Property Let Nombre(newV As String)
    mNombre = newV
End Property

Public Property Get NombreBase() As String
    NombreBase = mNombreBase
End Property

Public Property Let NombreBase(newV As String)
    mNombreBase = newV
End Property

Public Property Get ModoQueSeCompra() As ModoCompra
    ModoQueSeCompra = mModoQueSeCompra
End Property

Public Property Let ModoQueSeCompra(newV As ModoCompra)
    mModoQueSeCompra = newV
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property

Public Property Let Descripcion(newV As String)
    mDescripcion = newV
End Property

Public Property Get Publico() As Long
    Publico = mPublico
End Property

Public Property Let Publico(newV As Long)
    mPublico = newV
End Property

Public Property Get PrecioActual() As Long
    PrecioActual = mPrecioActual
End Property

Public Property Let PrecioActual(newV As Long)
    mPrecioActual = newV
End Property

Public Property Get PrecioActualVIP() As Long
    PrecioActualVIP = mPrecioActualVIP
End Property

Public Property Let PrecioActualVIP(newV As Long)
    mPrecioActualVIP = newV
End Property

Public Property Get CantidadPedida() As Long
    CantidadPedida = mCantidadPedida
End Property

Public Property Let CantidadPedida(newV As Long)
    mCantidadPedida = newV
End Property

Public Property Get HelpUse() As String
    HelpUse = mHelpUse
End Property

Public Property Let HelpUse(newV As String)
    mHelpUse = newV
End Property

