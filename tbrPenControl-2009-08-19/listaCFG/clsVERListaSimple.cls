VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVERListaSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'lista de opciones a la derecha

Private mOpciones() As String 'lista de opciones reales, valores internos que no se muestran, son como tags
Private mOpciones2() As String 'lista de opciones usadas (las visibles)

Private selOp As String 'opcion elegida
Private SelOpInternal As String 'elegida pero la escondida

'necesito una opcion temporal mientras el cliente manosea. Solo cuando acepte pasa a ser la oficial
Private selOpTMP As String 'opcion elegida temporal
Private SelOpInternalTMP As String 'elegida pero la escondida temporal

Private FSo As New Scripting.FileSystemObject

'la lista debe contener 2 elementos
Public Function LoadFromString(listaSeparada As String, separador As String, Optional separador2 As String = "") As Long
    'si separador2="" entonces es una lista simple y chota sin tags
    If separador2 = "" Then
        mOpciones = Split(listaSeparada, separador)
        mOpciones2 = mOpciones 'por compatibilidad con 2 listas
    Else
        'hay un caso especial cuando la lista son PATH y solo quiero mostrar lo ultimo
        Dim sp() As String, sp2() As String, K As Long
        
        If separador2 = "PATHS" Then
            sp = Split(listaSeparada, separador)
            Dim soloBase As String
            For K = 0 To UBound(sp)
                soloBase = FSo.GetBaseName(sp(K))
                'voy asignando a cada parte lo que corresponde
                ReDim Preserve mOpciones(K): ReDim Preserve mOpciones2(K)
                mOpciones(K) = sp(K): mOpciones2(K) = soloBase
            Next K
        Else
            sp = Split(listaSeparada, separador)
            For K = 0 To UBound(sp)
                If sp(K) <> "" Then
                    sp2 = Split(Trim(sp(K)), separador2)
                    'voy asignando a cada parte lo que corresponde
                    ReDim Preserve mOpciones(K): ReDim Preserve mOpciones2(K)
                    mOpciones2(K) = Trim(sp2(0)): mOpciones(K) = Trim(sp2(1))
                End If
            Next K
        End If
    End If
    
    LoadFromString = UBound(mOpciones)
End Function

Public Function GetStringListVisible() As String()
    GetStringListVisible = mOpciones2
End Function

Public Function GetStringListInternos() As String()
    GetStringListInternos = mOpciones
End Function

Public Function GetOpVisible(i As Long)
    GetOpVisible = mOpciones2(i)
End Function

Public Function GetOpTag(i As Long)
    GetOpTag = mOpciones(i)
End Function

Public Function GetSelectOp() As String
    GetSelectOp = selOp
End Function

Public Function GetSelectOpInternal() As String
    GetSelectOpInternal = SelOpInternal
End Function

Public Function GetSelectOpTMP() As String
    GetSelectOpTMP = selOpTMP
End Function

Public Function GetSelectOpInternalTMP() As String
    GetSelectOpInternalTMP = SelOpInternalTMP
End Function

'trata de elegir un elemento de la lista
Public Function TryToSelectFromInternalOptions(sel As String) As String
'como se usa para buscar paths y cosas no exactas usa instr
'pero por ejemplo para la tecla ok donde el codigo es 13 debe ser exacto y encuentra el 113 (teclaF2)
'asi que busco en toda la lista si hay uno identico toma ese sobre el "parecido"
    Dim K As Long
    For K = 0 To UBound(mOpciones)
    
        If LCase(mOpciones(K)) = LCase(sel) Then 'busca en las opciones internas IDENTICAS
            selOpTMP = mOpciones2(K)
            SelOpInternalTMP = mOpciones(K)
            TryToSelectFromInternalOptions = mOpciones(K)
            Exit Function
        End If
    Next K
    
    'SI NO ENCONTRO NADA IGUAL BUSCA ALGO PARECIDO
    For K = 0 To UBound(mOpciones)
        If InStr(LCase(mOpciones(K)), LCase(sel)) Then 'busca en las opciones PARECIDAS (sirve para paths)
            selOpTMP = mOpciones2(K)
            SelOpInternalTMP = mOpciones(K)
            TryToSelectFromInternalOptions = mOpciones(K)
            Exit Function
        End If
    Next K
    
    TryToSelect = "No encontrado"
End Function

Public Function DefineSelectFromID(sID As Long)
    If sID > UBound(mOpciones2) Then
        terr.AppendLog "daswq", CStr(sID) + "/" + CStr(UBound(mOpciones2))
    Else
        selOpTMP = mOpciones2(sID)
        SelOpInternalTMP = mOpciones(sID)
    End If
End Function

'trata de elegir un elemento de la lista
Public Function TryToSelectFromVisibleOptions(sel As String) As String
    Dim K As Long
    
    For K = 0 To UBound(mOpciones2)
        If LCase(mOpciones2(K)) = LCase(sel) Then 'busca en las opciones visibles IGUALES!!
            selOpTMP = mOpciones2(K)
            SelOpInternalTMP = mOpciones(K)
            TryToSelectFromVisibleOptions = mOpciones2(K)
            Exit Function
        End If
    Next K
    
    For K = 0 To UBound(mOpciones2)
        If InStr(LCase(mOpciones2(K)), LCase(sel)) Then 'busca en las opciones visibles PARECIDAS!!
            selOpTMP = mOpciones2(K)
            SelOpInternalTMP = mOpciones(K)
            TryToSelectFromVisibleOptions = mOpciones2(K)
            Exit Function
        End If
    Next K
    
    TryToSelect = "No encontrado"
End Function

Private Sub Class_Initialize()
    selOp = "NULL" 'marca para avisar que esta virgen y se maneja con selOpTmp
    SelOpInternal = "NULL"
    selOpTMP = ""
    SelOpInternalTMP = ""
    ReDim mOpciones(0) 'evita errores en getstriglist si se usa al iniciar mal
    ReDim mOpciones2(0) 'evita errores en getstriglist si se usa al iniciar mal
End Sub

Public Sub ConfirmOption()
    selOp = selOpTMP
    SelOpInternal = SelOpInternalTMP
End Sub
