VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuickLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clase para acceder a las traducciones exportadas
'es la clase liviana que solo contiene un idioma

Public Type QuickTrad
    sId As String
    sText As String
End Type

Private mCadenas() As QuickTrad

Public Function Load(sFile As String) As Long
    
    Load = 1
    Clear
    If Fso.FileExists(sFile) = False Then Exit Function
    
    Dim TE As TextStream, R As String
    Set TE = Fso.OpenTextFile(sFile, ForReading, False)
        R = TE.ReadAll
    TE.Close
    
    Dim Sp() As String
    Dim SP2() As String
    
    Sp = Split(R, Chr(6))
    Dim J As Long
    For J = 0 To UBound(Sp)
        If Sp(J) <> "" Then
            SP2 = Split(Sp(J), Chr(5))
            Newcadena SP2(0), SP2(1)
        End If
    Next J
    
    Load = 0
    'txFinal(K) = txFinal(K) + PHRs(J).sID + Chr(5) + PHRs(J).GetTrans(RET(K)) + Chr(6)
End Function

Private Sub Class_Initialize()
    Clear
End Sub

Public Sub Clear()
    ReDim mCadenas(0)
End Sub

Private Sub Newcadena(newID As String, newText As String)
    Dim J As Long
    J = UBound(mCadenas) + 1
    ReDim Preserve mCadenas(J)
    mCadenas(J).sId = newID
    mCadenas(J).sText = newText
    
End Sub

Public Function GetText(ssId As String, Optional sVars As String = "") As String
    'sVars es la lista de valores de cada variable, el separador es el primer caracter que aparezca en sVaris
    'por ejemplo sVaris = "|3pm|tbrsoft|cordoba" donde el separador es "|" y %1% es 3pm ...

    Dim RET As String
    Dim J As Long
    
    RET = "" 'si no encuentro devuelve nulo
    For J = 1 To UBound(mCadenas)
        If LCase(mCadenas(J).sId) = LCase(ssId) Then
            RET = mCadenas(J).sText
            Exit For
        End If
    Next J
    
    If sVars <> "" Then
        Dim sep As String
        sep = Left(sVars, 1)
        
        Dim Sp() As String
        Sp = Split(sVars, sep)
        
        'reemplazar todo
        
        For J = 1 To UBound(Sp)
            RET = Replace(RET, "%" + CStr(J) + "%", Sp(J)) 'por ejemplo %2% es la variable 2
        Next J
    End If
    
    GetText = RET
End Function
