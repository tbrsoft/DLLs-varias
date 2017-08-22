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

'es por eso que diferencio los precio de lo que se reproduce de lo que se compra empaquetado

Private ListaTypos() As New clsTypo
Private Promos() As New clsPromos
Private mPrecioBase As Single
Private mCreditosBase As Long 'me gusta siempre 10
Private mCreditosBilletes As Long 'creditos de la segunda señal del monedero

Private tErr As New tbrErrores.clsTbrERR

Public Function AddTypo(mNombre As String, mDescripcion As String, ModoBuy As ModoCompra) As clsTypo
    tErr.Anotar "aaaa", mNombre, mDescripcion, ModoBuy
    Dim k As Long
    k = UBound(ListaTypos) + 1
    ReDim Preserve ListaTypos(k)
    
    ListaTypos(k).NombreBase = mNombre
    
    If ModoBuy = CompraPaqueteArmado Then ListaTypos(k).Nombre = mNombre + "Buy"
    If ModoBuy = PagaYDespuesUsa Then ListaTypos(k).Nombre = mNombre + "Play"
    
    ListaTypos(k).Descripcion = mDescripcion
    ListaTypos(k).ModoQueSeCompra = ModoBuy
    ListaTypos(k).Publico = 1 'obligatoriamente publico (si se creo desde alla se usa desde alla)
    
    Set AddTypo = ListaTypos(k) 'lo devuelve para que lo manoseen
End Function

Public Function GetTypoByName(sNombre As String) As clsTypo
    tErr.Anotar "aaab", sNombre
    Dim k As Long
    For k = 1 To UBound(ListaTypos)
        If LCase(ListaTypos(k).Nombre) = LCase(sNombre) Then
            tErr.Anotar "aaac"
            Set GetTypoByName = ListaTypos(k)
            Exit Function
        End If
    Next k
    tErr.AppendLog "aaad:" + sNombre, UBound(ListaTypos)
    Set GetTypoByName = ListaTypos(0) 'devuelve el vacio! (nombre="")
    tErr.AppendLog "NoTypo!! " + sNombre
End Function

'obtener todos los tipos de x modo
Public Function GetTyposByModo(Modo As ModoCompra) As clsTypo()
    tErr.Anotar "aaae", Modo
    Dim Res() As clsTypo
    ReDim Res(0)
    
    Dim v As Long
    For v = 1 To UBound(ListaTypos)
        tErr.Anotar "aaaf", v
        If ListaTypos(v).ModoQueSeCompra = Modo Then
            'hay algunos tipos genericos como solopaquetes y soloplay para englobar. No se supone que se devuelvan aqui
            If ListaTypos(v).Nombre <> "SoloPlay" And ListaTypos(v).Nombre <> "SoloPaquetes" Then
                Dim w As Long
                w = UBound(Res) + 1
                ReDim Preserve Res(w)
                Set Res(w) = ListaTypos(v)
                tErr.Anotar "aaag", Modo, w
            End If
        End If
    Next v
    tErr.Anotar "aaah"
    GetTyposByModo = Res
End Function

Public Function AddPromo(typo As clsTypo, XaCantidad As Long, PriceCred As Long) As clsPromos
    tErr.Anotar "aaai", XaCantidad, PriceCred, typo.Nombre
    'no hay validacion de que sea logico el precio que agregan 'SEGUIRAQUI
    'la agrego ordenada por que despues la necesito asi
    
    'no permito en cantidad o precio cero!
    If XaCantidad = 0 Or PriceCred = 0 Then
        AddPromo = Promos(0)
        tErr.AppendLog "aaaj", typo.Nombre
        Exit Function
    End If
    
    Dim k As Long, v As Long
    k = UBound(Promos)
    
    If k = 0 Then
        tErr.Anotar "aaak"
        ReDim Preserve Promos(1)
        Promos(1).TypoMM = typo
        Promos(1).Cantidad = XaCantidad
        Promos(1).Creditos = PriceCred
        
        Set AddPromo = Promos(k)
        tErr.Anotar "aaap4"
        Exit Function
    Else
        tErr.Anotar "aaal"
        For v = 1 To k
            'ver que tenga la misma cantidad y el mismo typo!
            If (Promos(v).Cantidad = XaCantidad) And (Promos(v).TypoMM.Nombre = typo.Nombre) Then 'si ya hay una promo para x cantidad es que estoy reemplazando NO AGREGANDO!!!
                
                Promos(v).Creditos = PriceCred
                Set AddPromo = Promos(v)
                tErr.Anotar "aaap3"
                Exit Function
                
            End If
            
            If Promos(v).Cantidad > XaCantidad Then 'el primero que supere es el espacio que debe ocupar para quedar ordenado
                ReDim Preserve Promos(k + 1) 'se crea uno nuevo solo si no estaba de antes (en ese caso solo lo cambia)
                Dim w As Long
                For w = k To v Step -1
                    tErr.Anotar "aaam", w, k, v
                    Promos(w + 1).TypoMM = Promos(w).TypoMM
                    Promos(w + 1).Cantidad = Promos(w).Cantidad
                    Promos(w + 1).Creditos = Promos(w).Creditos
                Next w
                tErr.Anotar "aaan", v
                'inserto el que corresponde
                Promos(v).TypoMM = typo
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
        Promos(k + 1).TypoMM = typo
        Promos(k + 1).Cantidad = XaCantidad
        Promos(k + 1).Creditos = PriceCred
        
        Set AddPromo = Promos(k + 1)
        tErr.Anotar "aaap"
        Exit Function
    End If
    
End Function

'elimino una promocion
'debe haber al menos una para cada typo, si es la ultima doy error
Public Function DeletePromo(typo As clsTypo, XaCantidad As Long) As Long
    Dim PRO() As clsPromos
    PRO = GetPromosByTypo(typo)
    
    If UBound(PRO) = 1 Then 'error por que es la ultima
        DeletePromo = 1
        Exit Function
    End If
    
    'tener en cuenta que las prmocociones son una gran matriz con todas juntas
            'se llama PROMOS() hay que ver que indice tiene esta en la general
            'si es la ultima recortar la matriz y si esta antes correr todo un lugar para atras
    Dim h As Long
    For h = 1 To UBound(Promos) 'sobre la matriz general sobre la que abri (solo para saber si hay una sola)
        
        If (Promos(h).TypoMM.Nombre = typo.Nombre) And (Promos(h).Cantidad = XaCantidad) Then   'eliminar esta!
            
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
Public Function GetPromosByTypo(typo As clsTypo) As clsPromos()
    On Local Error GoTo erGetPromo
    tErr.Anotar "aaaq", typo.Nombre
    Dim k As Long, ret() As clsPromos
    ReDim ret(0)
    For k = 1 To UBound(Promos)
        tErr.Anotar "aaar", k, typo.Nombre
        'ver si va
        Dim siVa As Boolean
        siVa = False
        
        If typo.Nombre = "Todos" Then siVa = True
        tErr.Anotar "aaar2", Promos(k).TypoMM.Nombre
        If Promos(k).TypoMM.Nombre = typo.Nombre Then siVa = True 'ver si es lo quiero!! SEGUIRAQUI
        tErr.Anotar "aaar3", Promos(k).TypoMM.ModoQueSeCompra
        
        If typo.Nombre = "SoloPaquetes" Then
            If Promos(k).TypoMM.ModoQueSeCompra = CompraPaqueteArmado Then siVa = True
        End If
        
        tErr.Anotar "aaar5"
        If typo.Nombre = "SoloPlay" Then
            If Promos(k).TypoMM.ModoQueSeCompra = PagaYDespuesUsa Then siVa = True
        End If
        
        tErr.Anotar "aaas", siVa
        If siVa Then
            Dim m As Long
            m = UBound(ret) + 1
            ReDim Preserve ret(m)
            Set ret(m) = Promos(k)
            tErr.Anotar "aaat"
        End If
    Next k
    tErr.Anotar "aaau"
    GetPromosByTypo = ret
    
    Exit Function
    
erGetPromo:
    tErr.AppendLog "erGetPromo", tErr.ErrToTXT(Err)
    Resume Next
End Function

'de cualquier typo que sea el programa necesitara saber cuanto cobrar cada vez que se usa o se compra
Public Function GetPrice(typo As clsTypo, ByVal XaCantidad As Long) As Single
    
    tErr.Anotar "aaav", typo.Nombre, XaCantidad
    'busca el precio para la cantidad indicada, si no existe igual baja hasta encontrar
    'por ejemplo si alguien compra 4 canciones y hay ofertas por 3 y por 5 deberia usarse el precio por 3
    
    If typo.ModoQueSeCompra = CompraPaqueteArmado Then
        tErr.Anotar "aaaw"
        'ver todas las promociones cual corresponde
        Dim PRs() As clsPromos
        PRs = GetPromosByTypo(typo)
        
        Dim v As Long, dif As Long, best As Long, sel As Long
        best = -1: sel = -1
        For v = 1 To UBound(PRs)
            dif = XaCantidad - PRs(v).Cantidad
            tErr.Anotar "aaax", v, dif
            If dif >= 0 And dif < best Then
                best = dif
                sel = v
            End If
        Next v
        tErr.Anotar "aaay", sel
        If sel = -1 Then
            GetPrice = 0
        Else
            GetPrice = Round((PRs(sel).Creditos / PRs(sel).Cantidad) * XaCantidad, 2)
        End If
    End If
    
    tErr.Anotar "aaaz"
    
    If typo.ModoQueSeCompra = PagaYDespuesUsa Then 'seguramente xacatidad=1
        'debe haber pasado un updatePricesPlay cada vez que cambio la cantidad de creditos
        GetPrice = typo.PrecioActual
        tErr.Anotar "aaba", typo.PrecioActual
    End If
    
End Function

'actualizar los precios de lo que se paga y despues se escucha de a uno
Public Function UpdatePricesPlay(VC As Single, ByRef Creditos As Single)
    tErr.Anotar "aabb", VC, Creditos
    'vc es variacion de creditos
    'creditos es la cantidad actual en el sistema (se puede cambiar dentro poniendo en cero)

    Dim T() As clsTypo, v As Long
    T = GetTyposByModo(PagaYDespuesUsa)
    tErr.Anotar "aabc"
    If VC <= 0 Then
        tErr.Anotar "aabd"
        'si se ejecutaron canciones o videos y los creditos llegan hasta un valor
        'menor de una cancion en la maxima oferta disponible
        'enonces el precio vuelve a lo normal
        If Creditos < GetMasBaratoPlay Then
            tErr.Anotar "aabe"
            Creditos = 0
        End If
        
        'ahora pasar uno por uno de los tipo play y ver si tiene que cambiar precio
        For v = 1 To UBound(T)
            tErr.Anotar "aabf", v
            'tanto si entra plata como si sale si todo esta en cero es por que esta en modo gratis
            If UBound(GetPromosByTypo(T(v))) = 0 Then
                tErr.Anotar "aabg"
                T(v).PrecioActual = 0
            Else
                tErr.Anotar "aabh"
                If Creditos < GetMasBarato(T(v)) Then
                    If GetPrimerPrecio(T(v)) > 0 Then
                        T(v).PrecioActual = GetPrimerPrecio(T(v))
                    Else
                        tErr.Anotar "aabi"
                        T(v).PrecioActual = 1000000 'si no entra en ninguno ponemos precio inalcanzable
                    End If
                End If
            End If
        Next v
    End If
    tErr.Anotar "aabj"
    'si se pusieron monedas entonces el precio puede cambiar
    If VC > 0 Then
        tErr.Anotar "aabk"
        For v = 1 To UBound(T)
            'tanto si entra plata como si sale si todo esta en cero es por que esta en modo gratis
            If UBound(GetPromosByTypo(T(v))) = 0 Then
                tErr.Anotar "aabl"
                T(v).PrecioActual = 0
            Else
                tErr.Anotar "aabm"
                'dar vuelta por todas las promociones desde cantidad una en adelante, siempre esta ordenada por cantidad de menor a mayor, por eso se puede usar la funcion asi
                Dim P() As clsPromos, w As Long
                P = GetPromosByTypo(T(v))
                For w = 1 To UBound(P)
                    tErr.Anotar "aabn", w
                    If Creditos >= P(w).Creditos And P(w).Creditos > 0 Then 'no deberia ser nunca cero ya que no tiene sentido insertar un precio en cero, salvo los gratuitos
                        tErr.Anotar "aabo"
                        T(v).PrecioActual = tbrFIX(Round(P(w).Creditos / 2, 4), 2)
                        '(porque son los creditos xa 2 canciones)
                    End If
                Next w
            End If
        Next v
    End If

    tErr.Anotar "aabp"
    
    'habia una negrada que decia que se hiciera la siguiente verificacion que ahora no entiendo pero la respeto
    For v = 1 To UBound(T)
        If UBound(GetPromosByTypo(T(v))) > 0 Then
            tErr.Anotar "aabq"
            If T(v).PrecioActual = 0 Then
                tErr.Anotar "aabr", T(v).Nombre
                T(v).PrecioActual = 1000000
            End If
        End If
    Next v
    tErr.Anotar "aabs"
End Function

'necesario para calculo de precios y creditos
Private Function GetMasBarato(typo As clsTypo) As Single
    tErr.Anotar "aabt", typo.Nombre
    Dim PRs() As clsPromos
    PRs = GetPromosByTypo(typo)
    
    Dim v As Long
    Dim minP As Single: minP = -1
    For v = 1 To UBound(PRs)
        tErr.Anotar "aabu", v
        If (minP < 0) Or (minP < (PRs(v).Creditos / PRs(v).Cantidad)) Then
            tErr.Anotar "aabv"
            minP = (PRs(v).Creditos / PRs(v).Cantidad)
        End If
    Next v
    tErr.Anotar "aabw", minP
    GetMasBarato = minP
End Function

'necesario para calculo de precios y creditos
Private Function GetPrimerPrecio(typo As clsTypo) As Single
    tErr.Anotar "aabx", typo.Nombre
    'vendria a ser el mas caro, en realidad cuando uno se gasta una promocion debe volver al precio base
    'o sea al primero (por una unidad generalmente) que deberia ser el mas caro
    
    Dim PRs() As clsPromos
    PRs = GetPromosByTypo(typo)
    
    Dim v As Long
    Dim minCant As Single: minCant = -1
    Dim FinalP As Long
    
    For v = 1 To UBound(PRs)
        tErr.Anotar "aaby", v
        If (minCant < 0) Or (minCant < (PRs(v).Cantidad)) Then
            minCant = PRs(v).Cantidad
            FinalP = PRs(v).Creditos
        End If
    Next v
    
    tErr.Anotar "aabz", FinalP
    GetPrimerPrecio = FinalP
End Function

Private Function GetMasBaratoPlay() As Single
    tErr.Anotar "aaca"
    Dim T As clsTypo
    T = GetTypoByName("SoloPlay")
    
    GetMasBaratoPlay = GetMasBarato(T)
End Function

Private Sub Class_Initialize()
    tErr.FileLog = App.Path + "log.cs"
    tErr.Anotar "aacb"
    Clear
End Sub

Public Sub SetPathError(sFile As String)
    tErr.FileLog = sFile
End Sub

'es un string que dice algo "mostrable"
Public Function GetStrPreciosPlay(lFormat As Long, Separador As String) As String
    'lformat = 0 es plata ($)
    '        = 1 creditos
    
    'separador es por que el lblPrecios 2 es horizantal (sep = /) y el otro es vertical (sep = vbcrlf)
    
    tErr.Anotar "aacd", lFormat, Separador
    Dim tmp As String
    tmp = ""
    
    Dim T() As clsTypo, v As Long
    T = GetTyposByModo(PagaYDespuesUsa)
    Dim P() As clsPromos
    Dim w As Long
    'pasar por cada tipo de play y poner sus promociones
    For v = 1 To UBound(T)
        tErr.Anotar "aace", v
        
        'tanto si entra plata como si sale si todo esta en cero es por que esta en modo gratis
        P = GetPromosByTypo(T(v))
        
        Dim isGratis As Boolean
        If UBound(P) = 0 Then isGratis = True
        If UBound(P) = 1 Then
            If P(1).Creditos = 0 Then isGratis = True
        End If
        
        If isGratis Then
            tErr.Anotar "aacf"
            tmp = tmp + T(v).NombreBase + " gratis" + Separador
        Else
            tErr.Anotar "aacg"
            P = GetPromosByTypo(T(v))
            For w = 1 To UBound(P)
                tErr.Anotar "aach", w
                If lFormat = 0 Then tmp = tmp + T(v).NombreBase + CStr(P(w).Cantidad) + " x $" + CStr(P(w).Creditos * mPrecioBase) + Separador
                If lFormat = 1 Then tmp = tmp + T(v).NombreBase + CStr(P(w).Cantidad) + " x " + CStr(P(w).Creditos) + " créditos" + Separador
            Next w
        End If
        
        tErr.Anotar "aaci"
        Dim includeVIP As Boolean 'incluir los vip si los tienen en mas de cero
        includeVIP = (T(v).PrecioActualVIP > 0)
        
        'ver si esta habilitado cancion VIP
        If includeVIP Then
            tErr.Anotar "aacj"
            If lFormat = 0 Then tmp = tmp + T(v).NombreBase + " VIP x $" + CStr(T(v).PrecioActualVIP * mPrecioBase) + Separador
            If lFormat = 1 Then tmp = tmp + T(v).NombreBase + " VIP x $" + CStr(T(v).PrecioActualVIP) + " créditos" + Separador
        End If
    Next v
    
    tErr.Anotar "aack", tmp
    GetPrecios = tmp
End Function

Public Function GetStrPrecioPaquete(lFormat As Long, Separator As String)
    
    Dim h As Long, h2 As Long, tmp As String
    Dim T() As tbrPrecioRK.clsTypo
    T = PRI.GetTyposByModo(CompraPaqueteArmado)

    Dim PRO() As tbrPrecioRK.clsPromos
    
    For h = 1 To UBound(T)
        PRO = PRI.GetPromosByTypo(T(h))
        For h2 = 1 To UBound(PRO)
            If PRO(h2).Cantidad = 1 Then
                If lFormat = 0 Then tmp = tmp + "1 fichero de " + T(h).NombreBase + " por $ " + CStr(Round((PRO(h2).Creditos) * mPrecioBase, 2)) + Separador
                If lFormat = 1 Then tmp = tmp + "1 fichero de " + T(h).NombreBase + " por $ " + CStr(PRO(h2).Creditos) + Separador
            Else
                If lFormat = 0 Then _
                    tmp = tmp + CStr(h) + " ficheros de " + T(h).NombreBase + _
                        " por $ " + CStr(Round(PRO(h2).Creditos * mPrecioBase, 2)) + _
                        " ($ " + CStr(Round((PRO(h2).Creditos / PRO(h2).Cantidad) * mPrecioBase, 2)) + " cada uno)" + Separator
                If lFormat = 1 Then _
                    tmp = tmp + CStr(h) + " ficheros de " + T(h).NombreBase + _
                        " por $ " + CStr(PRO(h2).Creditos) + _
                        " ($ " + CStr(Round((PRO(h2).Creditos / PRO(h2).Cantidad), 2)) + " cada uno)" + Separator
            End If
        Next h2
    Next h
End Function

Public Property Get PrecioBase() As Single
    PrecioBase = mPrecioBase
End Property

Public Property Let PrecioBase(newV As Single)
    mPrecioBase = newV
End Property

Public Property Get CreditosBase() As Long
    CreditosBase = mCreditosBase
End Property

Public Property Let CreditosBase(newV As Long)
    mCreditosBase = newV
End Property

Public Property Get CreditosBilletes() As Long
    CreditosBilletes = mCreditosBilletes
End Property

Public Property Let CreditosBilletes(newV As Long)
    mCreditosBilletes = newV
End Property

Public Function Clear()
    
    mPrecioBase = 0
    mCreditosBase = 0
    mCreditosBilletes = 0
    
    ReDim Promos(0) 'para evitar errores y no usar el valor cero
    'varios tipos predefinidos para busquedas tipicas
    ReDim ListaTypos(3) 'para evitar errores y no usar el valor cero
    
    ListaTypos(1).Nombre = "Todos" 'el nombre es fijo, no tocarlo!
    ListaTypos(1).Descripcion = "Todos los tipos de contenido"
    ListaTypos(1).ModoQueSeCompra = Todos
    ListaTypos(1).Publico = 0
    
    ListaTypos(2).Nombre = "SoloPaquetes" 'el nombre es fijo, no tocarlo!
    ListaTypos(2).Descripcion = "Todos el contenido que se vende en paquetes"
    ListaTypos(2).ModoQueSeCompra = CompraPaqueteArmado
    ListaTypos(2).Publico = 0
    
    ListaTypos(3).Nombre = "SoloPlay" 'el nombre es fijo, no tocarlo!
    ListaTypos(3).Descripcion = "Todos el contenido que se reproduce"
    ListaTypos(3).ModoQueSeCompra = PagaYDespuesUsa
    ListaTypos(3).Publico = 0
    tErr.Anotar "aacc"
End Function

Public Function LoadPrices(sFile As String)
    
    On Local Error GoTo errLOAD
    
    
    Dim FSO As New Scripting.FileSystemObject
    
    If FSO.FileExists(sFile) = False Then
        LoadPrices = 1
        Exit Function
    End If
    
    'vaciar todo
    Clear
    
    Dim TE As TextStream, rr As String
    Set TE = FSO.OpenTextFile(sFile, ForReading)
        rr = TE.ReadAll
    TE.Close
    
    'los datos importantes vienen separados en chr(6)
    Dim SP5() As String 'separador de campos internos
    Dim SP6() As String 'separador de la base de esta cfg y de cada tipo de mm
    Dim SP7() As String 'separador de cada promocion
    
    Dim k As Long, m As Long, n As Long
    
    SP6 = Split(rr, Chr(6)) 'tengo un paquete con cada config basica + cada tipo y cada grupo de promociones de cada tipo
    
    Dim T As clsTypo
    Dim P As clsPromos
    
    For k = 0 To UBound(SP6)
        'datos internos de cada seccion
        If SP6(k) <> "" Then
            SP5 = Split(SP6(k), Chr(5))
        Else
            ReDim SP5(0)
            SP5(0) = ""
        End If
        
        Select Case SP5(0)
            Case "prMode" 'primera parte con las bases
                tErr.AppendSinHist SP5(1) 'registro de que version inicio
                
            Case "PrecioBase"
                mPrecioBase = CSng(SP5(1))
                If mPrecioBase = 0 Then mPrecioBase = 1
                
            Case "Creditosbase"
                mCreditosBase = CSng(SP5(1))
                If mCreditosBase = 0 Then mCreditosBase = 1
                
            Case "CreditosBilletes"
                mCreditosBilletes = CSng(SP5(1))
            
            'definicion de un tipo de multimedia
            Case "prTypo"
                Set T = AddTypo(SP5(1), SP5(2), CLng(SP5(3)))
                    T.Publico = CLng(SP5(4))
                    T.PrecioActualVIP = CLng(SP5(5))
                    'solo pòr seguridad, todos deberian tener al menos uno
                    If UBound(SP5) >= 7 Then
                        Set P = AddPromo(T, CLng(SP5(6)), CLng(SP5(7))) 'una promocion es obligatoria
                        For n = 8 To UBound(SP5) Step 2
                            Set P = AddPromo(T, CLng(SP5(n)), CLng(SP5(n + 1))) 'mas de una promocion es opcional
                        Next n
                    Else
                        tErr.AppendLog "noPromo8271:" + T.Nombre, SP6(k)
                    End If
        End Select
    Next k
    Exit Function
errLOAD:
    tErr.AppendLog "aact", tErr.ErrToTXT(Err)
    LoadPrices = Err.Number
End Function

Public Function SavePrices(sFile As String) As Long
    On Local Error GoTo errSave
    
    'grabar todos los datos de esta dll en un archivo
    tErr.Anotar "aacl"
    Dim FSO As New Scripting.FileSystemObject
    If FSO.FileExists(sFile) Then FSO.DeleteFile sFile, True
    
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(sFile, True)
        tErr.Anotar "aacm"
        TE.Write "prMode" + Chr(5) + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + Chr(6)
        TE.Write "PrecioBase" + Chr(5) + CStr(mPrecioBase) + Chr(6)
        TE.Write "Creditosbase" + Chr(5) + CStr(mCreditosBase) + Chr(6)
        TE.Write "CreditosBilletes" + Chr(5) + CStr(mCreditosBilletes) + Chr(6)
        
        tErr.Anotar "aacn"
        Dim k As Long, T As clsTypo
        For k = 4 To UBound(ListaTypos) 'hay tipo que engloban otros fijos en las posiciones 1 2 3, no tiene promociones propias
            TE.Write "prTypo" + Chr(5)
            Set T = ListaTypos(k)
            tErr.Anotar "aaco", k, T.Nombre, T.ModoQueSeCompra
            TE.Write _
                T.NombreBase + Chr(5) + _
                T.Descripcion + Chr(5) + _
                CStr(T.ModoQueSeCompra) + Chr(5) + _
                CStr(T.Publico) + Chr(5) + _
                CStr(T.PrecioActualVIP) + Chr(5) 'corto en 5 por que obligatoriamente tendra un precio y va pegadop como parte basica del tipo
            
                'ahora la lista de todas las promos de este tipo
                Dim v As Long, PR() As clsPromos
                PR = GetPromosByTypo(T)
                For v = 1 To UBound(PR)
                    tErr.Anotar "aacp", v, PR(v).Cantidad, PR(v).Creditos
                    TE.Write CStr(PR(v).Cantidad) + Chr(5) + CStr(PR(v).Creditos)
                    If v < UBound(PR) Then TE.Write Chr(5)
                Next v
             
            TE.Write Chr(6) 'separador de cada tipo
        Next k
        tErr.Anotar "aacq"
    TE.Close
    
    tErr.Anotar "aacr"
    SavePrices = 0
    Exit Function
errSave:
    tErr.AppendLog "aacs", tErr.ErrToTXT(Err)
    SavePrices = Err.Number
End Function

'lista con los nombres de los tipos
Public Function GetTypos(segunModo As ModoCompra) As String()
    Dim m As Long, ret() As String, siVa As Boolean
    ReDim ret(0)
    For m = 1 To UBound(ListaTypos)
        
        siVa = False
        If segunModo = Todos Then siVa = True
        If ListaTypos(m).ModoQueSeCompra = segunModo Then siVa = True
        
        If siVa Then
            ReDim Preserve ret(m)
            ret(m) = ListaTypos(m).Nombre
        End If
    Next m
    
    GetTypos = ret
End Function

Public Function ResetCantidadPedida()
    Dim m As Long
    For m = 1 To UBound(ListaTypos)
        ListaTypos(m).CantidadPedida = 0
    Next m
End Function

Public Function PESOSBase() As Single
    PESOSBase = Round(mPrecioBase / mCreditosBase, 2)
End Function

Public Function GetPesosPrecioActual(sNombreTypo As String) As Single
    GetPesosPrecioActual = GetTypoByName(sNombreTypo).PrecioActual * PESOSBase
End Function

Public Function GetPesosPrecioActualVIP(sNombreTypo As String) As Single
    GetPesosPrecioActualVIP = GetTypoByName(sNombreTypo).PrecioActualVIP * PESOSBase
End Function
