VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "modPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clase para manejar precios de rokoclas
'algunas cosas se meten al carrito y despues se compran, _
  es facil ya que se busca la promocion mas cercana y listo _
  pero cuando es musica o videos puedo tener por ejemplo 5 canciones por $ 1 y 3 videos por $ 1 _
  entonces el tipo mete $ 1, ejecuta una cancion, se la descuento $ 0,20 y depues mete un video !?

Private Promos() As New clsPromos

'por fuera yo llamo a esta clase para usarse como precios de ejecucion (lista de promocioes) y su derivado "precioActual" que es un calculo complejo
'tambien para precios VIP y muy importante para precios del expendedor por descargas

Private mTipoPrecio As TipoPrecios

'si es nueva la agrega y si ya existe la modifica
Public Function SetPromo(XaCantidad As Long, PriceCred As Long) As clsPromos
    tErr.Anotar "aaai", XaCantidad, PriceCred
    'no hay validacion de que sea logico el precio que agregan 'SEGUIRAQUI
    'la agrego ordenada por que despues la necesito asi
    
    'no permito en cantidad o precio cero!
    If XaCantidad = 0 Or PriceCred = 0 Then
        Set AddPromo = Promos(0)
        'a veces se pasa 1,0 y es para indicar que no hay vips
        'por lo tanto desactivo como error
        'tErr.AppendLog "aaaj:", CStr(XaCantidad) + "/" + CStr(PriceCred)
        Exit Function
    End If
    
    Dim k As Long, v As Long
    k = UBound(Promos)
    
    If k = 0 Then 'es nueva
        tErr.Anotar "aaak"
        ReDim Preserve Promos(1)
        Promos(1).Cantidad = XaCantidad
        Promos(1).Creditos = PriceCred
        
        Set AddPromo = Promos(k)
        tErr.Anotar "aaap4"
        Exit Function
    Else
        tErr.Anotar "aaal"
        For v = 1 To k
            'ver que tenga la misma cantidad
            If (Promos(v).Cantidad = XaCantidad) Then  'si ya hay una promo para x cantidad es que estoy reemplazando NO AGREGANDO!!!
                
                Promos(v).Creditos = PriceCred
                Set AddPromo = Promos(v)
                tErr.Anotar "aaap3"
                Exit Function
                
            End If
            
            'hago esto para que queden ordenados ya que en algun otro lado es necesario
            If Promos(v).Cantidad > XaCantidad Then 'el primero que supere es el espacio que debe ocupar para quedar ordenado
                ReDim Preserve Promos(k + 1) 'se crea uno nuevo solo si no estaba de antes (en ese caso solo lo cambia)
                Dim W As Long
                For W = k To v Step -1
                    tErr.Anotar "aaam", W, k, v
                    Promos(W + 1).Cantidad = Promos(W).Cantidad
                    Promos(W + 1).Creditos = Promos(W).Creditos
                Next W
                tErr.Anotar "aaan", v
                'inserto el que corresponde
                Promos(v).Cantidad = XaCantidad
                Promos(v).Creditos = PriceCred
                
                Set AddPromo = Promos(v)
                tErr.Anotar "aaap2"
                Exit Function
            End If
            tErr.Anotar "aaao"
        Next v
        
        'si llego hasta aqui es por que XaCant es mas grande que todos y por lo tanto va al ultimo
        ReDim Preserve Promos(k + 1)
        'inserto el que corresponde
        Promos(k + 1).Cantidad = XaCantidad
        Promos(k + 1).Creditos = PriceCred
        
        Set AddPromo = Promos(k + 1)
        tErr.Anotar "aaap"
        Exit Function
    End If
    
End Function

'elimino una promocion
'debe haber al menos una para cada typo, si es la ultima doy error
Public Function DeletePromo(XaCantidad As Long) As Long
    
    If UBound(Promos) = 1 Then 'error por que es la ultima
        DeletePromo = 1
        Exit Function
    End If
    
    'tener en cuenta que las prmocociones son una gran matriz con todas juntas
            'se llama PROMOS() hay que ver que indice tiene esta en la general
            'si es la ultima recortar la matriz y si esta antes correr todo un lugar para atras
    Dim h As Long
    For h = 1 To UBound(Promos) 'sobre la matriz general sobre la que abri (solo para saber si hay una sola)
        
        If (Promos(h).Cantidad = XaCantidad) Then   'eliminar esta!
            
            If h = UBound(Promos) Then
                
                ReDim Preserve Promos(h - 1)
                DeletePromo = 0 'se encontro y se borro
                Exit Function
                
            Else
                
                Dim h2 As Long
                For h2 = h To UBound(Promos) - 1
                    Set Promos(h2) = Promos(h2 + 1)
                Next h2
                
                h2 = UBound(Promos) - 1 'no me acuerdo si el next final me lo chongea
                ReDim Preserve Promos(h2) 'le recorte el ultimo
                
                DeletePromo = 0 'se encontro y se borro
                Exit Function
                
            End If
            
        End If
    Next h
    
    DeletePromo = 2 'no la encontre!
    
End Function

'devuelve todas las promociones de un tipo de contenido, si se pasa sin parametros van todas
'estan siempre ordenadas de menor a mayor cantidad (desde el addpromo)
Public Function GetPromos() As clsPromos()
    GetPromos = Promos
End Function

Public Sub AsignPromos(p() As clsPromos)
    p = Promos
End Sub


Public Function GetPromo(i As Long) As clsPromos
    If i = 0 Or i > UBound(Promos) Then
        Set GetPromo = Promos(0)
    Else
        Set GetPromo = Promos(i)
    End If
End Function


'esta funcion no es simple, tiene que ver con comprender integramente como funciona el sistema de precios
'lo mejor es no tocarla !

'por ejemplo precioActual que es un kilombo calcular cada vez que se calcula se asigna
'este actual nunca tendra promos, siempre es uno. Al igual que vip

'en caso de expendedor hay por ejemplo promociones por 1, 5 o 10 elementos y el tipo viene y compra 6
'en ese caso debe calcularse como base el precio por 5 y sacarle precio unitario de eso y multiplicarlo por 6

'tambien estan las promociones por reproduccion que son la base del calculo de precio actual

Public Function GetPrice(Optional ByVal XaCantidad As Long = 1, Optional UseSCM As Boolean = True) As Single
    
    tErr.Anotar "aaav", XaCantidad
    'si tengo la cantidad que el tipo pide se la doy, si no la calculo (solo pasara en expendedor)
    'ver todas las promociones cual corresponde
    
    Dim v As Long, dif As Long, best As Long, sel As Long
    best = -1: sel = -1
    For v = 1 To UBound(Promos)
        dif = XaCantidad - Promos(v).Cantidad
        tErr.Anotar "aaax", v, dif
        If dif >= 0 And dif < best Then
            best = dif
            sel = v
        End If
    Next v
    
    tErr.Anotar "aaay", sel
    If sel = -1 Then
        GetPrice = 0 'no encontro ni bosta !
    Else
        Dim VF As Single
        VF = Round((Promos(sel).Creditos / Promos(sel).Cantidad) * XaCantidad, 2) 'hasta aqui sabe el valor en creditos
        
        If UseSCM And ShowCreditsMode = 0 Then
            VF = Round(VF * Pesos1Cred, 2) 'ahora el modulo de precios sabe de plata a traves de la negrada del modulo
        End If
        
        GetPrice = VF
    End If

    
    tErr.Anotar "aaaz"

End Function

'se supone que este es un precio del tipo "precioActual" y depende de valores de promociones en otro objeto de este mismo tipo
Public Function UpdatePrice(BasePrecios As modPrecios, VC As Single, ByVal Creditos As Single) As Long

    'devulevo un long ya que se puede detectar aqui restos de creditos minimos
    'que deben ser llevados a cero. Como esa desicion se toma mas arriba yo solo indico
    'si ya no hay precios posible para el crédito disponible

    'valores que devuelve
    'UpdatePrice = 0 'bandera de permiso de poner los creditos oficiales en cero
    'UpdatePrice = 1 'bandera de que NO SE DEBE poner en cero todos los creditos
    'UpdatePrice = 2 'bandera de que NO TENGO NI IDEA si se deben poner en cero
    
    tErr.Anotar "aabc", VC, Creditos
    
    'tanto si entra plata como si sale si todo esta en cero es por que esta en modo gratis
    If BasePrecios.GetPrimerPrecio = 0 Then
        tErr.Anotar "aabg"
        SetPromo 1, 0 'pone el precio en cero ya que esto esta programado como gratis !
        UpdatePrice = 2 'bandera de que NO TENGO NI IDEA si se deben poner en cero los creditos
        Exit Function
    End If
    
    If VC <= 0 Then
        tErr.Anotar "aabf", VC, Creditos
        
        If Creditos < BasePrecios.GetMasBarato Then 'si ya no alcanza ni para ni voy al precio basico (si existe)
            'si se ejecutaron canciones o videos y los creditos llegan hasta un valor
            'menor de una cancion en la maxima oferta disponible
            'enonces el precio vuelve a lo normal
            
            Dim PP As Long
            PP = BasePrecios.GetPrimerPrecio
            tErr.Anotar "aabf2", PP
            SetPromo 1, PP
            tErr.Anotar "aabe"
            UpdatePrice = 0 'bandera de permiso de poner los creditos oficiales en cero
        Else
            'todavia puede comprar este !!
            UpdatePrice = 1 'bandera de que NO SE DEBE poner en cero todos los creditos
        End If

    End If
    
    tErr.Anotar "aabj", VC, Creditos
    'si se pusieron monedas entonces el precio puede cambiar
    If VC > 0 Then
        tErr.Anotar "aabk"
        
        'dar vuelta por todas las promociones desde cantidad una en adelante,
        'siempre esta ordenada por cantidad de menor a mayor, por eso se puede usar la funcion asi
        Dim W As Long
        For W = 1 To UBound(BasePrecios.GetPromos)
            tErr.Anotar "aabn", W
            If Creditos >= BasePrecios.GetPromo(W).Creditos And BasePrecios.GetPromo(W).Creditos > 0 Then  'no deberia ser nunca cero ya que no tiene sentido insertar un precio en cero, salvo los gratuitos
                tErr.Anotar "aabo"
                SetPromo 1, tbrFIX(Round(BasePrecios.GetPromo(W).Creditos / 2, 4), 2)
                '(porque son los creditos xa 2 canciones)
                UpdatePrice = 1 'bandera de que NO SE DEBE poner en cero todos los creditos
            End If
        Next W
        
    End If

    tErr.Anotar "aabp"
    
    'habia una negrada que decia que se hiciera la siguiente verificacion que ahora no entiendo pero la respeto
    
    'quiere decir algo como que si hay precios y de todas formas vos estas en cero entonces andate al carajo
    If UBound(BasePrecios.GetPromos) > 0 Then
        tErr.Anotar "aabq"
        If GetPrice(1) = 0 Then
            tErr.Anotar "aabr"
            SetPromo 1, 1000000
        End If
    End If

    tErr.Anotar "aabs"

End Function

'necesario para calculo de precios y creditos
'SIEMPRE EN CREDITOS, NO EN $
Public Function GetMasBarato() As Single
    tErr.Anotar "aabt"
    
    Dim v As Long
    Dim minP As Single: minP = -1
    For v = 1 To UBound(Promos)
        tErr.Anotar "aabu", v
        If (minP < 0) Or (minP < (Promos(v).Creditos / Promos(v).Cantidad)) Then
            tErr.Anotar "aabv"
            minP = Round((Promos(v).Creditos / Promos(v).Cantidad), 2)
        End If
    Next v
    tErr.Anotar "aabw", minP
    'no requiere pasarse a pesos ya que es para calculos internos en base a creditos
    GetMasBarato = minP
End Function

'necesario para calculo de precios y creditos
'SIEMPRE EN CREDITOS, NO EN $
Public Function GetPrimerPrecio() As Long
    tErr.Anotar "aabx"
    'vendria a ser el mas caro, en realidad cuando uno se gasta una promocion debe volver al precio base
    'o sea al primero (por una unidad generalmente) que deberia ser el mas caro si se programo bien
    
    Dim v As Long
    Dim minCant As Single: minCant = -1
    Dim FinalP As Long
    
    For v = 1 To UBound(Promos)
        tErr.Anotar "aaby", v
        If (minCant < 0) Or (minCant > (Promos(v).Cantidad)) Then
            minCant = Promos(v).Cantidad
            FinalP = Promos(v).Creditos
        End If
    Next v
    
    tErr.Anotar "aabz", FinalP
    GetPrimerPrecio = FinalP
End Function

Private Sub Class_Initialize()
    tErr.Anotar "aacb"
    Clear
End Sub

Public Function Clear()
    ReDim Promos(0) 'para evitar errores y no usar el valor cero
    tErr.Anotar "aacc"
End Function

'aqui solo se devuelven las cantidades y los precios
'este precio no sabe quien, solo es un PromosManager

'si este objeto es un VIP y no esta activado entonces habra una solo
'promo y costara cero, si pasa los mismo con la musica significa
'que el usuario desea pasar musica gratis. para no confundir esto
'devuelvo "" cuando es gratis, de el otro la do el sistema sabra que hacer
Public Function GetStrPrecios(Separador As String, Optional UseSCM As Boolean = True) As String
    
    'separador es por que el lblPrecios 2 es horizantal (sep = /) y el otro es vertical (sep = vbcrlf)

    tErr.Anotar "aacd", lFormat, Separador
    Dim tmp As String
    tmp = ""

    Dim W As Long
    'pasar por cada tipo de play y poner sus promociones
    For v = 1 To UBound(Promos)
        tErr.Anotar "aace", v

        'tanto si entra plata como si sale si todo esta en cero es por que esta en modo gratis

        Dim isGratis As Boolean
        If UBound(Promos) = 0 Then isGratis = True
        If UBound(Promos) = 1 Then
            If Promos(1).Creditos = 0 Then isGratis = True
        End If

        If isGratis Then
            tErr.Anotar "aacf"
            tmp = "" 'del otro lado sabran que hacer si da "", sirve para que la traduccion no invada esta clase
        Else
            tErr.Anotar "aach", v
            If UseSCM And ShowCreditsMode = 0 Then
                tmp = tmp + CStr(Promos(v).Cantidad) + " x $" + CStr(Round(Promos(v).Creditos * Pesos1Cred, 2)) + Separador
            Else
                tmp = tmp + CStr(Promos(v).Cantidad) + " x " + CStr(Round(Promos(v).Creditos, 2)) + " cred." + Separador
            End If
        End If
    Next v

    tErr.Anotar "aack", tmp
    GetStrPrecios = tmp
End Function

Public Property Get TipoPrecio() As TipoPrecios
    TipoPrecio = mTipoPrecio
End Property

Public Property Let TipoPrecio(newVal As TipoPrecios)
    mTipoPrecio = newVal
End Property
