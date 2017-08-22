VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTypoMM_mng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Dim FSO As New Scripting.FileSystemObject
Private ListaTypos() As New clsTypoMM  'todas las existentes

'********************************************
'PRECIOS
Private mPrecioBase As Single
Private mCreditosBase As Single
Private mCreditosBilletes As Single
'********************************************


Public Function AddTypo(sID As MMTypes, mNombre As String) As clsTypoMM
    tErr.Anotar "aaaa", mNombre, mDescripcion
    Dim k As Long
    k = UBound(ListaTypos) + 1
    ReDim Preserve ListaTypos(k)
    
    ListaTypos(k).IdTypo = sID
    ListaTypos(k).Nombre = mNombre ' es algo mas mio, traducible y modificable
    
    Set AddTypo = ListaTypos(k) 'lo devuelve para que lo manoseen
End Function

Public Function GetTypoByID(sID As Long) As clsTypoMM
    tErr.Anotar "aaab09", sNombre
    If sID = 0 Or sID > UBound(ListaTypos) Then
        Set GetTypoByID = ListaTypos(0) 'devuelve el vacio! (nombre="")
        tErr.AppendLog "NoTypo!!(8) " + sNombre
    Else
        Set GetTypoByID = ListaTypos(sID) 'devuelve el vacio! (nombre="")
    End If
End Function


Public Function GetTypo(sID As MMTypes) As clsTypoMM
    tErr.Anotar "aaab", sNombre
    Dim k As Long
    For k = 1 To UBound(ListaTypos)
        If ListaTypos(k).IdTypo = sID Then
            tErr.Anotar "aaac"
            Set GetTypo = ListaTypos(k)
            Exit Function
        End If
    Next k
    tErr.AppendLog "aaad:" + sNombre, UBound(ListaTypos)
    Set GetTypo = ListaTypos(0) 'devuelve el vacio! (nombre="")
    tErr.AppendLog "NoTypo!! " + sNombre
End Function

Private Function GetNumId_ByTypo(sID As MMTypes) As Long
    tErr.Anotar "aaab5", sNombre
    Dim k As Long
    For k = 1 To UBound(ListaTypos)
        If ListaTypos(k).IdTypo = sID Then
            tErr.Anotar "aaac5"
            GetNumId_ByTypo = k
            Exit Function
        End If
    Next k
    tErr.AppendLog "aaad5:" + sNombre, UBound(ListaTypos)
    GetNumId_ByTypo = 0 'devuelve el vacio! (nombre="")
    tErr.AppendLog "NoTypo5!!" + sNombre
End Function


Public Sub SetPathError(sFile As String)
    tErr.FileLog = sFile
End Sub

Private Sub Class_Initialize()
    tErr.FileLog = App.Path + "log.cs"
    tErr.Anotar "aacb"
    Clear
    mCreditosBase = 1 'para que no divida por cero
End Sub

'seguiraqui deberia ser posible que trabaje mas rapido ,si el tipo no activo la venta de extras
'si va a ser predeterminado la venta de otros contenidos esta bien asi
Public Function ObtenerArchMM2(Carpeta As String, ByRef TipoFinal As MMTypes) As String()
    
    'devuelve una matriz solo con los paths de los archivos elegidos
    'devuelve ademas el tipo final elegido
    'si typofinal es algun valor ya definido lo respeto y uso ese sin volver a calcular
    'si typoFinal es cero entonces lo busco yo
    
    'la propiedad ACTIVO de los typos me dice si se usaran o NO (NUEVO AGO-09)
    
    On Error GoTo ErrObtMM
    
    tErr.Anotar "001-0214", Carpeta
    If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta + "\"
    
    'aqui viene los resultados
    Dim TMPmatriz() As String
    ReDim Preserve TMPmatriz(0)
    
    'una vez cargado esto por fuera se define un perfil del disco para ver de que tipo es
    'tambie estaría bueno definir un origen ya con características de tipo de disco
    'por que no solo son disco multimedia si no que hay
    'discos de aplicaciones JAVA
    'discos de ringtones (ideal para confundirse con de mp3)
    'discos de wallapers
    
    'deberia dar un perfil a cada disco, si por ejemplo hay muchos mp3 de _
        mas de 2 minutos y aparece un mp3 de 30 segundos y 2 o 3 imagenes _
        deberia darse cuenta que es un disco comun _
        Si la mayoria son imagenes es una carpeta de wallpapers _
        Si la mayoria son mp3 de menos de 2 minutos es una carpeta de ringtones _
        Si hay 10 jar y 10 jpg debo interpretar que es una carpeta de java con _
        sus screenshots correspondientes
        
    'para esto hay un dato que sirve, una cancion MP3 con calidad comun ocupa 1 MB por cada minuto
    'POR ejemplo en el disco "En vivo en Cemento (10-10-1998) - A morir !!! de catupecu _
        hay 4 canciones de menos de 1.5 MB pero hay 18 MP3s de mas de 1.5 MB _
        el sistema debería identificar este disco como de musica !
    
    'SEGUIRAQUI sacar las comas a los nombre del los archivos es una paquetada
    'corregir para que esto no sea un problema !!
    
    Dim FO As Folder
    Dim FI As File
    Set FO = FSO.GetFolder(Carpeta)
    
    Dim ContTP() As Long 'cuenta cuantos de cada tipo hay
    ReDim ContTP(UBound(ListaTypos)) 'una por cada typo existente
    
    'voy viendo cual es el que tiene mas y cual es el indice del typo ganador
    Dim MaxCant As Long, MaxPos As Long
    MaxCant = -1: MaxPos = -1
    
    'ya voy pasando todo a una matriz para devolver (de esta matriz luego se depurara los que realmente van segun el tipo de disco)
    Dim tmpMTX() As String
    ReDim tmpMTX(0)
    'cada elemento de la matriz esta asignado a un solo tipo de MM (salvo excepciones)
    'en el caso de un disco con 10 mp3s largos y uno corto este ultimo se validara
    'como ringtones y el disco finalmenmte ganara el tipo "mp3-musica comun" y el de ringtones
    'puede quedar afuera, en este caso cada elemento debe tener varias opciones y si alguna gana se incluye
        
    Dim m As Long 'usos varios
    
    For Each FI In FO.Files
        
        tErr.Anotar "001-0217", FI.Name
        
        'revisar este archivo contra todos los tipos posibles
        Dim Pos As Long, Otros As String
        ValidateFilebyTipo FI.Path, Pos, Otros
        
        'agregarlo a la matriz resultante
        m = UBound(tmpMTX) + 1
        ReDim Preserve tmpMTX(m)
        If Otros = "" Then
            
            tmpMTX(m) = CStr(Pos) + Chr(5) + FI.Path 'despues se podran sacar solo los de un tipo para el resultado final
            ContTP(Pos) = ContTP(Pos) + 2 'dos puntos al pulenta y uno al secundario
            tErr.Anotar "001-0218", Pos, ContTP(Pos)
            'tengo que llevar un contador de cuantas son de cada tipo para definir el tipo del disco
            If ContTP(Pos) > MaxCant Then
                MaxCant = ContTP(Pos)
                MaxPos = Pos
            End If
            
        Else 'es un ringtone o bicho raro
            
            'en la lista de otros viene el ganador tambien
            tmpMTX(m) = Otros + Chr(5) + FI.Path
            
            Dim SP6() As String, n As Long
            SP6 = Split(Otros)
            For n = 0 To UBound(SP6)
                tErr.Anotar "001-0219", SP6(n), ContTP(CLng(SP6(n)))
                'dos puntos al pulenta y uno al secundario
                If CLng(SP6(n)) = Pos Then
                    ContTP(CLng(SP6(n))) = ContTP(CLng(SP6(n))) + 2
                Else
                    ContTP(CLng(SP6(n))) = ContTP(CLng(SP6(n))) + 1
                End If
                
                'tengo que llevar un contador de cuantas son de cada tipo para definir el tipo del disco
                If ContTP(CLng(SP6(n))) > MaxCant Then
                    MaxCant = ContTP(CLng(SP6(n)))
                    MaxPos = CLng(SP6(n))
                End If
                
            Next n
        End If
            
    Next
    
    tErr.Anotar "001-0220", MaxPos, MaxCant, ListaTypos(Pos).IdTypo, TipoFinal
    
    '*****************************************************************
    If TipoFinal = NADA Then 'AUTODETECTAR el que gano
    '*****************************************************************
        TipoFinal = ListaTypos(MaxPos).IdTypo
        'pos se definio ya, no hay que tocarlos
    '*****************************************************************
    Else 'USAR EL QUE SE PIDIO
    '*****************************************************************
        'se ignora el que gano y se usa el que venia como parametro
        Pos = GetNumId_ByTypo(TipoFinal)
        'la posicion pos no se sabe, se averigua mediante la funcion usada
    End If
    '*****************************************************************
    
    
    'el tipo autodetectado es ListaTypos(MaxPos)
    'obtener entonces la matriz resultante
    Dim MTXFINAL() As String
    ReDim Preserve MTXFINAL(0)
    
    Dim h As Long, Sp() As String
    For h = 1 To UBound(tmpMTX)
        tErr.Anotar "001-0221", h, tmpMTX(h)
        
        Sp = Split(tmpMTX(h), Chr(5))
        
        Dim SP5() As String
        SP5 = Split(Sp(0))
        
        'si cualquiera califica PONERLO!
        For m = 0 To UBound(SP5)
            
            tErr.Anotar "001-022", m, SP5(m), MaxPos
            
            If CLng(SP5(m)) = MaxPos Then
                m = UBound(MTXFINAL) + 1
                ReDim Preserve MTXFINAL(m)
                MTXFINAL(m) = Sp(1)
            End If
            
        Next m
    Next h
    
    ObtenerArchMM2 = MTXFINAL
    
    Exit Function
ErrObtMM:
    tErr.AppendLog tErr.ErrToTXT(Err), "Archivos.bas" + ".acpk4"
    Resume Next
    
End Function

'revisa el archivo dado y lo controla contra todos los tipos existentes, devuelve el tipo que corresponde
'PUEDE SER QUE CORRESPONDA A MAS DE UN TIPO:
'en el caso de un disco con 10 mp3s largos y uno corto este ultimo se validara
'como ringtone y el disco finalmenmte ganara el tipo "mp3-musica comun" y el de ringtones
'puede quedar afuera, en este caso cada elemento debe tener varias opciones y si alguna otra gana se incluya tambien

'los ISOs son un caso parecido, si bien parece más lógico mostrar ISOs de CD y de DVD en el mismo
'disco si voy a mostrar los precios de cada tipo en dentro de cada disco
'lo mejor sera separarlos. En este caso NO quiero que un tipo pase a ser del otro
'como SI lo quiero en ringtones vs MP3 comunes

'son los unicos dos casos en que tengo conflicto de extenciones pero en cada uno
'necesito una resolucion diferente ... ???

'a cada tipo le agrego una propiedad que indique si este tipo tambien puede ser interpretado como otro
'donde el valor predet sera falso y solo en ringtone sera true (en MP3-comun sera falso)

Private Function ValidateFilebyTipo(ByVal sFile As String, ReturnIX As Long, Optional ByRef OderWin As String) As clsTypoMM
    'sFile es el archivo a analizar
    'ReturnIX es el ganador
    'OtrosWin son otros tipos que tambien aplicarian por la extencion. Son los IDs de typo separado por espacios
    
    OderWin = ""
    Dim h As Long, tmp As Long
    
    Dim MaxVal As Long, MaxPos As Long 'valor maximo y posicion donde esta
    MaxVal = -1: MaxPos = -1
    
    
    For h = 1 To UBound(ListaTypos)
        'la propiedad ACTIVO de los typos me dice si se usaran o NO (NUEVO AGO-09)
        If ListaTypos(h).Activo > 0 Then
            tmp = ListaTypos(h).Validate(sFile) 'revisa por extencion y si hay dos que la cumplen "urga" un poco mas
            
            If tmp > 0 And ListaTypos(h).PuedeSerOtro Then
                If OderWin <> "" Then OderWin = OderWin + " "
                OderWin = OderWin + CStr(i)
            End If
            
            If tmp > MaxVal Then
                MaxVal = tmp
                MaxPos = h
            End If
        End If
    Next h
    
    If MaxPos = -1 Then 'no es de ningun tipo conocido (basura en las carpetas o los "*.3H" mios
        Set ValidateFilebyTipo = ListaTypos(0)
        ReturnIX = 0
    Else
        ReturnIX = MaxPos
        Set ValidateFilebyTipo = ListaTypos(MaxPos)
    End If
    
End Function

'parecida a la anterior pero ya con valores para usar desde afuera mas practicos
Public Function ValidateTypo(sFile As String) As MMTypes
    Dim IX As Long, OW As String
    ValidateTypo = ValidateFilebyTipo(sFile, IX, OW).IdTypo
End Function

'*******************************************************************************************
'precios
'*******************************************************************************************

'actualizar los precios de lo que se paga y despues se escucha de a uno

'si el tipo puso $1 por la promocion de 3 canciones por $2 a medida que va agregando
'monedas el precio va cambiando
'asi tambien cuando use las 3 (el precio de cada una era 0.66) quedaran $ 0.02
'que deben ser eliminados

Public Function UpdatePrices(VC As Single, ByRef Creditos As Single) As Long
    
    tErr.Anotar "aabb", VC, Creditos
    'vc es variacion de creditos (cada vez que se cambia esto se recalcula el precio actual de reproduccion)
    'creditos es la cantidad actual en el sistema (se puede cambiar dentro poniendo en cero cuando corresponda)

    Dim J As Long, PonerCeroCred As Long, lastVal As Long, Vals(2) As Long
    
    For J = 1 To UBound(ListaTypos)
        If ListaTypos(J).CanPlay Then
            lastVal = ListaTypos(J).PR.UpdatePriceBase(VC, Creditos)
            Vals(lastVal) = Vals(lastVal) + 1
        End If
    Next J
    
    'con que alguno sea uno alcanza para no llevarlo a cero
    'si alguno es cero y ninguno es uno se debe poner en cero
    If Vals(1) > 0 Then
        PonerCeroCred = 0
    Else
        If Vals(0) Then
            PonerCeroCred = 1
            Creditos = 0 'YA LO PONGO AQUI EN CERO
        End If
    End If
    
    UpdatePrices = PonerCeroCred
    
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
    Dim SP6() As String 'separador de la base de esta cfg y de cada tipo de mm
    Dim SP7() As String 'separador de cada promocion
    Dim SP5() As String 'separador de campos internos

    Dim k As Long, m As Long, n As Long

    SP6 = Split(rr, Chr(6)) 'tengo un paquete con cada config basica + cada tipo y cada grupo de promociones de cada tipo

    Dim T As clsTypoMM 'el ultimo tipo elegido y el considera actual para la carga

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
                Set T = AddTypo(CLng(SP5(1)), SP5(3))
                T.Activo = CLng(SP5(2))
                T.Descripcion = SP5(4)
                'la variable T se "acuerda" del ultimo elegido para embocar los precios
            Case "P1"
                Dim c As Long
                c = LoadOnePrice(T.PR.PricePlay, SP6(k))
                'solo pòr seguridad, todos deberian tener al menos uno
                If c < 2 Then tErr.AppendLog "noPromo8271:" + T.Nombre, SP6(k)
            
            Case "P2"
                Dim c2 As Long
                c2 = LoadOnePrice(T.PR.PriceDescarga, SP6(k))
                'solo por seguridad, todos deberian tener al menos uno
                If c2 < 2 Then
                    tErr.AppendLog "noPromo8272:" + T.Nombre, SP6(k)
                    T.PR.PriceDescarga.SetPromo 1, 1
                End If
            
            Case "P3"
                Dim c3 As Long
                c3 = LoadOnePrice(T.PR.PriceVIP, SP6(k))
                'este puede no tener promos
                If c3 < 2 Then
                    'T.PR.PriceDescarga.SetPromo 1, 0
                    'la promo en cero es rechazada por precios, asi que para saber que no esta activado vip
                    'debere preguntar si no hay promos
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
        Dim k As Long, T As clsTypoMM
        For k = 1 To UBound(ListaTypos) 'hay tipo que engloban otros fijos en las posiciones 1 2 3, no tiene promociones propias
            TE.Write "prTypo" + Chr(5)
            Set T = ListaTypos(k)
            tErr.Anotar "aaco", k, T.Nombre
            TE.Write _
                CStr(T.IdTypo) + Chr(5) + _
                CStr(T.Activo) + Chr(5) + _
                T.Nombre + Chr(5) + _
                T.Descripcion + Chr(6)
            
            'escribir todos los precios de este. El precio actual se calcula en base a precios play
            TE.Write "P1" + Chr(5) + StrSavePrice(ListaTypos(k).PR.PricePlay) + Chr(6)
            TE.Write "P2" + Chr(5) + StrSavePrice(ListaTypos(k).PR.PriceDescarga) + Chr(6)
            TE.Write "P3" + Chr(5) + StrSavePrice(ListaTypos(k).PR.PriceVIP) + Chr(6)
                
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

Private Function StrSavePrice(p As modPrecios) As String
    'ahora la lista de todas las promos de este tipo
    Dim v As Long, ret As String
    ret = ""
    For v = 1 To UBound(p.GetPromos)
        tErr.Anotar "aacp", v, p.GetPromo(v).Cantidad, p.GetPromo(v).Creditos
        ret = ret + CStr(p.GetPromo(v).Cantidad) + Chr(5) + CStr(p.GetPromo(v).Creditos)
        If v < UBound(p.GetPromos) Then ret = ret + Chr(5)
    Next v
    
    StrSavePrice = ret
End Function

Private Function LoadOnePrice(p As modPrecios, sPrices As String) As Long
    'ahora la lista de todas las promos de este tipo
    Dim v As Long, Sp() As String
    Sp = Split(sPrices, Chr(5)) 'el indice cero dice P1, P2 ... no se usa aca
    
    For v = 2 To UBound(Sp) Step 2
        tErr.Anotar "aacp6", v, Sp(v), Sp(v - 1)
        p.SetPromo CLng(Sp(v - 1)), CLng(Sp(v))
    Next v
    
    LoadOnePrice = UBound(Sp)
End Function

'precio de una x cantidad de creditos de base (creditos base)
Public Property Get PR_PrecioBase() As Single
    PrecioBase = mPrecioBase
End Property

Public Property Let PR_PrecioBase(newV As Single)
    mPrecioBase = newV
    Pesos1Cred = Round(mPrecioBase / mCreditosBase, 2)
End Property

Public Property Get PR_CreditosBase() As Long
    CreditosBase = mCreditosBase
End Property

Public Property Let PR_CreditosBase(newV As Long)
    mCreditosBase = newV
    Pesos1Cred = Round(mPrecioBase / mCreditosBase, 2)
End Property

Public Property Get PR_CreditosBilletes() As Long
    CreditosBilletes = mCreditosBilletes
End Property

Public Property Let PR_CreditosBilletes(newV As Long)
    mCreditosBilletes = newV
End Property

Public Function PR_GetPesosBase1Cred() As Single
    PR_GetPesosBase1Cred = Pesos1Cred
End Function

Public Sub SetCreditsMode(newV As Long)
    ShowCreditsMode = newV
End Sub

'*******************************************************************************************
'CARRITO
'*******************************************************************************************

'reinicializar la cantidad de cosas pedidas para el carrito
Public Function Cart_ResetCantidadPedida()
    Dim m As Long
    For m = 1 To UBound(ListaTypos)
        ListaTypos(m).CantidadPedida = 0
    Next m
End Function

'saber cuanto cuesta todo lo que se pide para el carrito ahora
Public Function Cart_CalculateTotalPrice() As Single
        
    Dim toTal As Single
    toTal = 0
    
    For m = 1 To UBound(ListaTypos)
        toTal = toTal + ListaTypos(m).PR.PriceDescarga.GetPrice(ListaTypos(m).CantidadPedida)
    Next m
    
    Cart_CalculateTotalPrice = toTal
End Function

'*******************************************************************************************
'load, save and init
'*******************************************************************************************

Public Sub Init()
    'luego de cargados todos los tipos son necesarias algunas revisiones
    
    '1- ver extenciones duplicadas para que el validador trabaje mas rapido o no
    Dim i As Long, J As Long
    Dim exts2() As Extes 'todas las extenciones usadas hasta el momento
    Dim totEX As Long 'contador de extenciones
    For i = 1 To UBound(ListaTypos)
        For J = 1 To ListaTypos(i).GetExtsCant
            totEX = totEX + 1
            ReDim Preserve exts2(totEX)
            exts2(totEX) = ListaTypos(i).GetExt(J)
            exts2(totEX).sName = LCase(exts2(totEX).sName)
            
            'revisar si alguna de las extenciones anteriores es esta
            Dim Rev As Long, k As Long
            Rev = totEX - 1
            For k = 1 To Rev
                If exts2(k).sName = exts2(totEX).sName Then
                    exts2(k).isDupli = True
                    exts2(totEX).isDupli = True
                End If
            Next k
            
        Next J
    Next i
End Sub

Public Function GetCantTipos() As Long
    GetCantTipos = UBound(ListaTypos)
End Function

Public Sub Clear()
    ReDim Preserve ListaTypos(0)
    ListaTypos(0).IdTypo = NADA
End Sub
