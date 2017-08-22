VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRHASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'esta clase esta compuesta por:

'el texto a traducir ...
Private mBaseText As String
'(que puede incluir variables dadas en el prog que usa esto)
'por ejemplo el texto puede decir
'"Software desarrollavo por %01% en %02%"
'donde en el soft se colocara %01% = tbrSoft and %02% = Argentina
'esto es para garantizar que el traductor no escriba boludeces

Private mVars() As clsVars 'todas las variables de esta frase

'despues de esto solo quedan las cadenas de texto finales en una cantidad variable de idiomas

Private mTrans() As clsTrans

'ademas tiene un id unico para que el software traducido pueda identificar univocamete
Private mnID As String 'los numeros simepre terminan siendo limitados

'la base del texto a veces se corrige y necesita re-traducirse
'para esto se guarda referencia a la fecha de la base y de cada traduccion
'las traducciones desactualzadas se mostraran al traductor responsable de cada proyecto

Private mFechaBase As Long

'tags para separar partes de la traduccion (gui, config, etc)
Private mTags() As String

'si el traductor cree que necesita decirle algo al que escribio el texto o pedirle ayuda
'deberia poder poner tags u oibservaciones


'listo todo definido ....

'-------------------------------------------
'PROPERTYs
Public Property Get BaseText() As String
    BaseText = mBaseText
End Property

Public Property Let BaseText(ByVal newV As String)
    'SI ES DIFERENTE A LA ANTERIOR SE ACTUALZA LA FECHA
    If Trim(LCase(newV)) <> Trim(LCase(mBaseText)) Then
        mFechaBase = CLng(Now)
    End If
    
    'si alguien usa una "{" o "}" me caga por que las uso para grabar !!!
    
    newV = Replace(newV, "{", "[")
    newV = Replace(newV, "}", "]")
    mBaseText = newV
    
End Property

Public Property Get sId() As String
    sId = mnID
End Property

Public Property Let sId(newV As String)
    mnID = newV
End Property

Public Function SetVar(num As Long, Expli As String) As clsVars 'si el numero de variable existela reemplaza y si o la crea
    'el mVars(0) es la explicacion de la base, luego los numeros de cada variable
    Dim J As Long
    J = UBound(mVars)
    If num > J Then
        ReDim Preserve mVars(num)
        Set mVars(num) = New clsVars
    End If
    J = num
    
    mVars(J).n = num 'simepre el indice es num
    mVars(J).mExplicacion = Expli
    
    Set SetVar = mVars(J)
End Function

Public Function GetVar(Optional n As Long = 0) As String 'el predeterminado es el cero por que se llamara para dar la explicacion del texto a traducir
    Dim RET As String
    If n > UBound(mVars) Then
        RET = ""
    Else
        RET = mVars(n).mExplicacion
    End If
    
    GetVar = RET
End Function

Public Function GetVarCantidad() As Long
    GetVarCantidad = UBound(mVars)
End Function

'setear una traduccion
'la fecha opcional es por que cuando se hace LOAD ya trae la fecha y hay que respetarla
Public Function SetTrans(ByVal txt As String, IDM As String, Optional Fecha As Long = -1) As clsTrans  'marcar una traduccion
    
    'si alguien usa una "{" o "}" me caga por que las uso para grabar !!!
    txt = Replace(txt, "{", "[")
    txt = Replace(txt, "}", "]")
    
    'buscar todas las traducciones por el mismo idioma
    Dim J As Long
    For J = 1 To UBound(mTrans)
        If LCase(mTrans(J).RefIdioma) = LCase(IDM) Then
        
            'SI ES DIFERENTE A LA ANTERIOR SE ACTUALZA LA FECHA
            If Trim(LCase(txt)) <> Trim(LCase(mTrans(J).Finaltext)) Then
                mTrans(J).fechaTrans = CLng(Now)
            End If
            
            mTrans(J).Finaltext = txt
            
            If Fecha <> -1 Then mTrans(J).fechaTrans = Fecha
            
            Set SetTrans = mTrans(J)
            Exit Function
        End If
    Next J
    
    'crear uno nuevo
    J = UBound(mTrans) + 1
    ReDim Preserve mTrans(J)
    Set mTrans(J) = New clsTrans
    mTrans(J).RefIdioma = IDM
    mTrans(J).Finaltext = txt
    If Fecha <> -1 Then
        mTrans(J).fechaTrans = Fecha
    Else
        mTrans(J).fechaTrans = CLng(Now)
    End If
    
    'quedeaqui VER QUE HACER CON TODAS LAS PROPIEDADES NUEVAS EN CASO DE QUE SEA NUEVO O EN "LOAD"
    Set SetTrans = mTrans(J)
    
End Function

'obtener el texto traducido en x idioma, si hay variables me las tiene que pasar
Public Function GetTrans(IDM As String, Optional sVaris As String = "") As String
    'svaris es la lista de valores de cada variable, el separador es el primer caracter que aparezca en sVaris
    'por ejemplo sVaris = "|3pm|tbrsoft|cordoba" donde el separador es "|" y %1% es 3pm ...

    Dim RET As String
    Dim J As Long
    
    RET = mBaseText 'si no encuentra el idioma quedara en la base
    For J = 1 To UBound(mTrans)
        If LCase(mTrans(J).RefIdioma) = LCase(IDM) Then
            RET = mTrans(J).Finaltext
            Exit For
        End If
    Next J
    
    
    If sVaris <> "" Then
        Dim sep As String
        sep = Left(sVaris, 1)
        
        Dim Sp() As String
        Sp = Split(sVaris, sep)
        
        'reemplazar todo
        
        For J = 1 To UBound(Sp)
            RET = Replace(RET, "%" + CStr(J) + "%", Sp(J)) 'por ejemplo %2% es la variable 2
        Next J
    End If
    
    GetTrans = RET
End Function

'--------------------------------------
'INIT
Private Sub Class_Initialize()
    Init
End Sub

Private Sub Init()
    ReDim mVars(0)
    Set mVars(0) = New clsVars
    mVars(0).n = 0
    mVars(0).mExplicacion = ""
    
    ReDim mTrans(0)
    Set mTrans(0) = New clsTrans
    
    ReDim mTags(0)
End Sub

'----------------------
'juntar splits separados!
Private Function JoinSPLIT(sSP() As String, ix As Long, sep As String) As String

    Dim RET As String
    RET = ""
    Dim J As Long
    For J = ix To UBound(sSP)
        RET = RET + sSP(J)
        If J < UBound(sSP) Then RET = RET + sep
    Next J
    
    JoinSPLIT = RET
    
End Function

'--------------------------------
'FILE
'toma todas las propiedaes y las empaqueta en un string
Public Function Save(Optional sFile As String = "") As String
    
    On Local Error GoTo ErrSave
    
    Dim RET As String
    
    'id unico
    RET = vbCrLf + vbCrLf + "{ID=" + mnID + "}" + vbCrLf
    
    'texto base
    RET = RET + "{baseText=" + mBaseText + "}" + vbCrLf
    RET = RET + "{fechaBase=" + CStr(mFechaBase) + "}" + vbCrLf
    
    'variables
    Dim J As Long
    For J = 0 To UBound(mVars) 'empiezo de cero por que mVars(0) es la explicacion del texto general, luego 1,2,3 ... de cada variable usada
        RET = RET + "{var=" + CStr(mVars(J).n) + "=" + mVars(J).mExplicacion + "}" + vbCrLf
    Next J
    
    For J = 1 To UBound(mTrans)
        RET = RET + "{trans=" + CStr(mTrans(J).fechaTrans) + _
            "=" + mTrans(J).RefIdioma + _
            "=" + Replace(mTrans(J).Finaltext, "=", "%IGUAL%") + _
            "=" + CStr(Abs(CLng(mTrans(J).Terminada))) + _
            "=" + CStr(Abs(CLng(mTrans(J).NoTerminada))) + _
            "=" + CStr(Abs(CLng(mTrans(J).NoEntiendo))) + _
            "=" + Replace(mTrans(J).Observaciones, "=", "%IGUAL%") + "}" + vbCrLf
            
    Next J
    
    For J = 1 To UBound(mTags)
        RET = RET + "{tags=" + CStr(mTags(J)) + "}" + vbCrLf
    Next J
    
    'opcionalmente se puede grabar a un archivo si especifica sFile
    If sFile <> "" Then
        Dim TE As TextStream
        Set TE = Fso.CreateTextFile(sFile, True)
            TE.Write RET
        TE.Close
    End If
    
    'devolver el resultado
    Save = RET
    
    Exit Function
ErrSave:
    TERR.AppendLog "ErrSavePhrase", TERR.ErrToTXT(Err)
    Save = ""
    
End Function

Public Function Load(sRead As String) As Long
    
    On Local Error GoTo ErrLoadPhrase
    
    'si sRead es un pat lo leo y si no es un texto a levantar a la clase
    If Fso.FileExists(sRead) Then
        Dim TE As TextStream
        Set TE = Fso.OpenTextFile(sRead, ForReading, False)
            sRead = TE.ReadAll
        TE.Close
    End If
    
    'buscar cada una de las cadenas a levantar entre {llaves}
    Dim Textos() As String 'cada una de las cadenas completas entre llaves
    Textos = GetLlaves(sRead)
    
    Dim J As Long, Sp() As String
    For J = 1 To UBound(Textos)
        Sp = Split(Textos(J), "=")
        
        Select Case LCase(Sp(0))
            Case "id"
                mnID = Trim(Sp(1))
            Case "basetext"
                mBaseText = Trim(Sp(1))
            Case "fechabase"
                mFechaBase = CLng(Trim(Sp(1)))
            Case "var"
                SetVar CLng(Sp(1)), JoinSPLIT(Sp, 2, "=") 'SP(2) podria tener "="s !
            Case "trans"
                Dim t As clsTrans
                Set t = SetTrans(Replace(Sp(3), "%IGUAL%", "="), Sp(2), CLng(Sp(1)))
                t.Terminada = CBool(Sp(4))
                t.NoTerminada = CBool(Sp(5))
                t.NoEntiendo = CBool(Sp(6))
                t.Observaciones = Replace(Sp(7), "%IGUAL%", "=")
                
            Case "tags"
                AddTag Sp(1)
            Case Else
                'anda a saber que sera!!!
                TERR.AppendSinHist "LoadExtraño:" + Textos(J)
        End Select
    
    Next J
    
    Load = 0
    
    Exit Function
ErrLoadPhrase:
    TERR.AppendLog "ErrLoadPhrase", TERR.ErrToTXT(Err)
    Load = Err.Number
    
End Function

'¿tiene el tag?
Public Function HasTag(sTag As String) As Boolean
    Dim J As Long, RET As Boolean
    RET = False
    For J = 1 To UBound(mTags)
        If LCase(mTags(J)) = LCase(sTag) Then
            RET = True
        End If
    Next J
    
    HasTag = RET
End Function

Public Sub AddTag(S As String)
    Dim J As Long
    J = UBound(mTags) + 1
    ReDim Preserve mTags(J)
    mTags(J) = S
End Sub

Public Sub RemoveTagByText(S As String)
    Dim J As Long, f As Long
    For J = 1 To UBound(mTags)
        If LCase(mTags(J)) = LCase(S) Then
            f = J
        End If
    Next J
    
    If f = UBound(mTags) Then
        ReDim Preserve mTags(f - 1)
    Else
        For J = f To UBound(mTags) - 1
            mTags(J) = mTags(J + 1)
        Next J
        f = UBound(mTags) - 1
        ReDim Preserve mTags(f)
    End If
End Sub

Public Function GetStrIdiomasByColon() As String
    Dim J As Long, RET As String
    For J = 1 To UBound(mTrans)
        RET = RET + mTrans(J).RefIdioma
        If J < UBound(mTrans) Then RET = RET + ";"
    Next J
    
    GetStrIdiomasByColon = RET
End Function

Public Function GetStrTagsByColon() As String
    Dim J As Long, RET As String
    For J = 1 To UBound(mTags)
        RET = RET + mTags(J)
        If J < UBound(mTags) Then RET = RET + ";"
    Next J
    
    GetStrTagsByColon = RET
End Function

'devuleve matriz de textos
Private Function GetLlaves(fullStr As String) As String()
    Dim RET() As String
    ReDim RET(0)
    
    Dim B As Long, B2 As Long
    Dim lastB As Long
    lastB = 1
    Do
        B = InStr(lastB, fullStr, "{")
        If B = 0 Then Exit Do 'se termino
        lastB = B + 1
        
        'debe haber una llave de cierre para ese
        B2 = InStr(lastB, fullStr, "}")
        If B2 = 0 Then Exit Do 'se termino MAL sin cierre, ignoro todo lo que sigue ?? SEGUIRAQUI
    
        'hay una cadena!
        Dim J As Long
        J = UBound(RET) + 1
        ReDim Preserve RET(J)
        RET(J) = Mid(fullStr, B + 1, B2 - B - 1)
    Loop
    
    GetLlaves = RET
    
End Function

Public Function GetTransObj(IDM As String) As clsTrans
    
    If IDM <> "" Then 'si no hay idioma se va mas abajo a lo predeterminado
        Dim J As Long
        For J = 1 To UBound(mTrans)
            If LCase(mTrans(J).RefIdioma) = LCase(IDM) Then
                Set GetTransObj = mTrans(J)
                Exit Function
            End If
        Next J
    End If
    
    Dim t As New clsTrans
    'si no lo encontro crea uno predeterminado BASICO
    t.RefIdioma = IDM 'si IDM no importa
    t.Finaltext = mBaseText 'NO TRADUCIDO, texto base
    t.NoTerminada = True 'quedeaqui asegurarme que lo predeterminado sea true
    t.Terminada = False
    
    Set GetTransObj = t
    
End Function

Public Property Get FechaBase() As Long
    FechaBase = mFechaBase
End Property

