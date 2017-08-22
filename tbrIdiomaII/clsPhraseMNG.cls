VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPhraseMNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'paquetes de phrases
'la usa el programador (para definir las cadenas) y el traductor (para traducir)
'debe exportarse algo simple ID-texto para que sea lo mas liviano posiblke que levanta el software

'para que los traductores no tengan que cargarse de tooooodo un software (podria ser lento buscar cada cadena)
'lo mejor es que se use un paquete de estos para cada formulario segun partes que se definan a gusto
'del programador

Private PHRs() As clsPRHASE
Private sFileActual As String 'si se hizo load o sabe grabo aqui donde fue por si lo quiero usar de nuevo


Public Function QuickAddPhr(ssId As String, sBaseText As String, Optional sVarsByPuntoyComa As String = "", Optional sTagsByPuntoyComa As String = "") As clsPRHASE
    'desde el codigo fuente del sistema necesito una forma muy rapida de agregar cadenas a traducir
    
    Dim p As New clsPRHASE, Sp() As String, J As Long
    p.sId = ssId
    p.BaseText = sBaseText
    
    If sVarsByComa <> "" Then
        'separarlapor coma e ir agregandolas
        Sp = Split(sVarsByComa, ";")
        For J = 0 To UBound(Sp)
            p.SetVar J, Sp(J)  'mVars(0) es la explicacion de la base
        Next J
    End If
    
    If sTagsByComa <> "" Then
        Sp = Split(sTagsByComa, ";")
        For J = 0 To UBound(Sp)
            p.AddTag Sp(J)
        Next J
    End If
    
    AppendPHR p
    Set QuickAddPhr = p
End Function

Public Sub AppendPHR(ByVal p As clsPRHASE)
    Dim J As Long
    J = UBound(PHRs) + 1
    ReDim Preserve PHRs(J)
    Set PHRs(J) = p
End Sub

Private Sub Class_Initialize()
    Clear
End Sub

Public Function getPHR(sId As String) As clsPRHASE
    Dim J As Long
    For J = 1 To UBound(PHRs)
        If LCase(sId) = LCase(PHRs(J).sId) Then
            Set getPHR = PHRs(J)
            Exit Function
        End If
    Next J
    
    'error ! no la encontro
    Set getPHR = PHRs(0)
End Function

Public Function GetPhrCantidad() As Long
    GetPhrCantidad = UBound(PHRs)
End Function

Public Function GetPhrByNum(n As Long) As clsPRHASE
    If n > 0 And n <= UBound(PHRs) Then
        Set GetPhrByNum = PHRs(n)
    Else
        'error, indice no valido!
        Set GetPhrByNum = PHRs(0)
    End If
End Function

Public Function Load(sRead As String)
    
    'si sRead es un pat lo leo y si no es un texto a levantar a la clase
    If Fso.FileExists(sRead) Then
        sFileActual = sRead
        Dim TE As TextStream
        Set TE = Fso.OpenTextFile(sRead, ForReading, False)
            sRead = TE.ReadAll
        TE.Close
        
    End If
    
    'limpiar todos los phrs
    Clear
    
    Dim Sp() As String
    Sp = Split(sRead, "&&&")
    
    Dim J As Long
    Dim p() As clsPRHASE
    For J = 0 To UBound(Sp)
        ReDim Preserve p(J)
        Set p(J) = New clsPRHASE 'si uso un solo objeto P se reasignan siempre al ultimo por mas que sea byVal el load
        p(J).Load (Sp(J))
        AppendPHR p(J)
    Next J
        
End Function

'devuelve un string y opcionalmnte lo graba  a un archivo
Public Function Save(Optional sFile As String = "") As String

    Dim RET As String
    Dim J As Long
    For J = 1 To UBound(PHRs)
    
        RET = RET + PHRs(J).Save
        If J < UBound(PHRs) Then RET = RET + "&&&"
    
    Next J

    
    If sFile = "LAST" Then 'quiere grabar el archivo que ya esta usando
        sFile = sFileActual
    End If
    
    'opcionalmente se puede grabar a un archivo si especifica sFile
    If sFile <> "" Then
        Dim TE As TextStream
        Set TE = Fso.CreateTextFile(sFile, True)
            TE.Write RET
        TE.Close
        sFileActual = sFile
    End If
    
    Save = RET
End Function

Public Sub Clear()
    ReDim PHRs(0)
End Sub

'una vez terminado necesito exportar archivos livianos uno con cada idioma (incluso la base de texto)
Public Function ExportAll(sFolder As String)
    If Fso.FolderExists(sFolder) = False Then Fso.CreateFolder (sFolder)
    
    'un archivo con cada idioma encontrado
    Dim RET() As String 'un string por cada idioma encontrado, para eso debo saber de antemano que idiomas existes y que cuando una cadena no este en un idioma dejar la base)
    'SEGUIRAQUI: por ejemplo si existe el aleman y no hay una cadena el ingles tiene prioridad, no deberia yo usar mi base en español que es mas lejana al aleman [QUEDA PARA DESPUES]
    
    RET = GetIdiomas
    
    'buscar todos los idiomas existentes
    Dim J As Long, K As Long
    
    Dim CantIdms As Long
    CantIdms = UBound(RET)
    
    Dim txFinal() As String 'cada uno de los textos completos en cada uno de los idiomas
    ReDim txFinal(CantIdms)
    
    For J = 1 To UBound(PHRs)
        'el idioma base tambien !!!
        txFinal(0) = txFinal(0) + PHRs(J).sId + Chr(5) + PHRs(J).BaseText + Chr(6)
        For K = 1 To CantIdms 'cada idioma es ret(k)
            txFinal(K) = txFinal(K) + PHRs(J).sId + Chr(5) + PHRs(J).GetTrans(RET(K)) + Chr(6)
        Next K
    Next J
    
    'escribirt los archivos a disco ...
    If Right(sFolder, 1) <> "\" Then sFolder = sFolder + "\"
    
    
    RET(0) = "Español" 'la base del texto tambien se graba!!!
    
    Dim sFile As String
    For K = 0 To CantIdms 'cada idioma es ret(k)
        sFile = Fso.BuildPath(sFolder, RET(K) + ".phr")
        Dim TE As TextStream
        Set TE = Fso.CreateTextFile(sFile, True)
            TE.Write txFinal(K)
        TE.Close
    
    Next K
    
End Function

Public Function GetIdiomas() As String()
    
    Dim RET() As String
    ReDim RET(0)
    'devuelve la lista de idiomas encontrados
    Dim J As Long, K As Long, L As Long, m As Long, Sp() As String, R As String
    For J = 1 To UBound(PHRs) 'ITERO POR CADA FRASE
        Sp = Split(PHRs(J).GetStrIdiomasByColon, ";")
        For K = 0 To UBound(Sp) 'ITERO POR CADA IDIOMA TRADUCIDO DE ESTA FRASE
            
            Dim Esta As Boolean
            Esta = False
            For L = 1 To UBound(RET) 'ITERO POR LOS IDIOMAS YA DETECTADOS PARA VER SI HAY ALGUNO QUE AGREGAR
                If LCase(RET(L)) = LCase(Sp(K)) Then
                    Esta = True
                    Exit For
                End If
            Next L
            
            If Esta = False Then
                m = UBound(RET) + 1
                ReDim Preserve RET(m)
                RET(m) = Sp(K) 'cada idioma queda en su indice SEGUIRAQUI, ver que el separador de lineas este ok para editarse y grabarse
            End If
            
        Next K
    Next J
    
    GetIdiomas = RET
End Function

'devuelve matriz con los tags sin repetir de toda esta coleccion
Public Function GetTags() As String()
    Dim RET() As String
    ReDim RET(0)
    'devuelve la lista de idiomas encontrados
    Dim J As Long, K As Long, L As Long, m As Long, Sp() As String, R As String
    For J = 1 To UBound(PHRs) 'ITERO POR CADA FRASE
        Sp = Split(PHRs(J).GetStrTagsByColon, ";")
        For K = 0 To UBound(Sp) 'ITERO POR CADA IDIOMA TRADUCIDO DE ESTA FRASE
            
            Dim Esta As Boolean
            Esta = False
            For L = 1 To UBound(RET) 'ITERO POR LOS IDIOMAS YA DETECTADOS PARA VER SI HAY ALGUNO QUE AGREGAR
                If LCase(RET(L)) = LCase(Sp(K)) Then
                    Esta = True
                End If
            Next L
            
            If Esta = False Then
                m = UBound(RET) + 1
                ReDim Preserve RET(m)
                RET(m) = Sp(K) 'cada idioma queda en su indice SEGUIRAQUI, ver que el separador de lineas este ok para editarse y grabarse
            End If
            
        Next K
    Next J
    
    GetTags = RET
End Function

'devuelve cantidad de cadenas traducidas ok
Public Function GetCadenasOk(IDM As String, Optional TAGs As String = "") As Long
    Dim RET As Long
    RET = 0
    
    Dim OkTag As Boolean
    
    For J = 1 To UBound(PHRs)
        
        If TAGs <> "" Then
            If PHRs(J).HasTag(TAGs) Then
                OkTag = True
            Else
                OkTag = False
            End If
        Else
            OkTag = True
        End If
        
        If OkTag Then
            If PHRs(J).GetTransObj(IDM).Terminada Then RET = RET + 1
        End If
        
    Next J
    
    GetCadenasOk = RET
End Function

Public Function GetCadenasCantidad(Optional sTag As String = "") As Long
    Dim RET As Long
    If sTag = "" Then
        RET = UBound(PHRs)
    Else
        RET = 0
        For J = 1 To UBound(PHRs)
            If PHRs(J).HasTag(sTag) Then RET = RET + 1
        Next J
    End If
    
    GetCadenasCantidad = RET
End Function

Public Function GetAutoID(Optional CantDigit As Long = 6) As String
    'devuelve un ID no usado con una cantidad predefinida de digitos
    
    Dim RET As String
    Dim S As String, J As Long
    For J = 1 To 10 * CantDigit
        S = String(CantDigit - Len(CStr(J)), "0") + CStr(J)
        If ExistsId(S) = False Then
            GetAutoID = S
            Exit Function
        End If
    Next J
    
End Function

Public Function ExistsId(ssId As String) As Boolean
    Dim J As Long
    For J = 1 To UBound(PHRs)
        If LCase(PHRs(J).sId) = LCase(ssId) Then
            ExistsId = True
            Exit Function
        End If
    Next J
    
    ExistsId = False
    
End Function
