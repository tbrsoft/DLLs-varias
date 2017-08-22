VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVARS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clase con las variables
'asi como && significa un salto de linea hay variables y se escriben dentro de %%

'llevo aqui una matriz de valores
Private Type myVar
    mVar As String 'nombre de la variable
    mVal As String 'valor
End Type

Private V() As myVar

'reemplaza todas las variables que haya en el sistema en un texto por los valores quer la fonola indique
Public Function TraslateVars(ByVal s As String) As String  'pasa siermpre por referencia para modificar el texto original
    terr.Anotar "qfa", s
    'SEGUIRAQUI que busque solo las variables que haya !!!
    'por ahora cada variable se agrega a mano. de esta forma seria multiuso la dll y no limitada a las variables
    Dim b As Long, b2 As Long, lastB As Long
    lastB = 1
    Do
        b = InStr(lastB, s, "%%")
        terr.Anotar "qfb", b
        If b > 0 Then
            terr.Anotar "qfc"
            b2 = InStr(b + 2, s, "%%")
            terr.Anotar "qfd", b2
            If b2 > 0 Then
                'por las dudas no dejo poner espacios de modo que no se confunda una variable con un %% por otro motivo
                Dim tVAR As String
                terr.Anotar "qfe", b, b2
                tVAR = mID(s, b + 2, b2 - b - 2) 'anda, lo que anda no se pregunta!
                'terr.AppendSinHist vbCrLf + vbCrLf + _
                    "***" + _
                    s + ":" + CStr(b) + ":" + CStr(b2) + vbCrLf + _
                    "VAR:" + tVAR + vbCrLf + _
                    "***" + vbCrLf
                FindVar s, tVAR
                
                lastB = b2 + 2 'para que siga buscando variables
                terr.Anotar "qff", lastB
            Else
                terr.Anotar "qfg"
                Exit Do
            End If
        Else
            terr.Anotar "qfh"
            Exit Do
        End If
        terr.Anotar "qfi"
    Loop
    
    terr.Anotar "qfj", s
    TraslateVars = s
    
End Function

'busca 1 o mas veces la variable para reemplazarla
Private Function FindVar(ByRef sTex As String, sVar As String) As Boolean
    terr.Anotar "qfk", sTex, sVar
    Do 'una o las veces que este3 la variable!
        If InStr(sTex, "%%" + sVar + "%%") Then
            terr.Anotar "qfl"
            sTex = Replace(sTex, "%%" + sVar + "%%", GetValueVar(sVar))
        Else
            terr.Anotar "qfm"
            Exit Do
        End If
        terr.Anotar "qfn"
    Loop
    terr.Anotar "qfo"
End Function

'obtiene del sistema el valor de la variable
Public Function GetValueVar(sVar As String) As String
    terr.Anotar "qfp", sVar
    Dim K As Long
    For K = 1 To UBound(V)
        terr.Anotar "qfr", K, V(K).mVar
        If LCase(V(K).mVar) = LCase(sVar) Then
            terr.Anotar "qfs"
            GetValueVar = V(K).mVal
            'terr.AppendSinHist "VAL-OK " + sVar + ":" + V(K).mVal
            Exit Function
        End If
        terr.Anotar "qft"
    Next K
    
    'terr.AppendSinHist "VAL-NO (" + CStr(UBound(V)) + ")" + sVar
    terr.Anotar "qfu"
    GetValueVar = "-1" 'NO ENCONTO VALOR!!
    
End Function


'////////////*/*/*/*/*/*/*/*
'////////////*/*/*/*/*/*/*/*
'Carga de variables
'////////////*/*/*/*/*/*/*/*
'////////////*/*/*/*/*/*/*/*
Public Sub SetV(sVar As String, sVal As String)  'cargar cualquier variable
    terr.Anotar "qfv", sVar, sVal
    Dim K As Long
    K = UBound(V) + 1
    ReDim Preserve V(K)
    V(K).mVar = sVar
    V(K).mVal = sVal
    terr.Anotar "qfw"
    'terr.AppendSinHist "NEWVAL " + sVar + ":" + sVal
End Sub

Private Sub Class_Initialize()
    ReDim V(0)
End Sub
