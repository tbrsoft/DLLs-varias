VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Load_Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'archivo del tipo LOAD que incluyen las claves de encriptacion y los indices usados para cargar licencias

Public sName As String
Public LL1 As String 'clave encriptacion [llaves de licencia de los archivos generados]
Public LL2 As String 'clave encr de licencia de este generador de licencias
Private ID1s(9, 5) As String 'indices de claves de licencias
Private ID2s(9, 5) As String 'opcional para manejar una segunda licencia (solo en la version desde el licenciero
Private ComoDevuelve(9) As Long 'segun que licencia del licenciero que tipo de licencia devuelve

Private fso As New Scripting.FileSystemObject

'cargar un archivo de datos de software que ya existe
Public Function Unload__D(sFile__D As String) As Long

    On Local Error GoTo ErrLoadPhrase
    Terr.Anotar "baf", sFile__D
    Unload__D = 1 'error
    
    
    Dim sFileBase As String
    sFileBase = sFile__D
    Dim sRead As String
    
    Dim Textos() As String 'cada una de las cadenas completas entre llaves
    Textos = GetLlaves(sFileBase)
    
    Dim Ret As New Load_Files
        
    Dim J As Long, SP() As String
    For J = 1 To UBound(Textos)
        If Textos(J) <> "" Then
            SP = Split(Textos(J), ":")
            Terr.Anotar "bag", Textos(J), SP(0)
            
            Select Case LCase(SP(0))
                Case "system"
                    sName = Trim(SP(1))
                Case "llave"
                    LL1 = Trim(SP(1))
                Case "llave2"
                    LL2 = Trim(SP(1))
                Case "setid"
                    Dim i1 As Long, i2 As Long, i3 As Long
                    i1 = CLng(Trim(SP(1)))
                    i2 = CLng(Trim(SP(2)))
                    i3 = CLng(Trim(SP(3)))
                    ID1s(i1, i2) = CStr(i3)
                    
                Case "setid2"
                    Dim j1 As Long, j2 As Long, J3 As Long
                    j1 = CLng(Trim(SP(1)))
                    j2 = CLng(Trim(SP(2)))
                    J3 = CLng(Trim(SP(3)))
                    ID2s(j1, j2) = CStr(J3)
                    
                Case "usenr"
                    ComoDevuelve(CLng(Trim(SP(1)))) = CLng(Trim(SP(2)))
                Case Else
                    'anda a saber que sera!!!
                    Terr.AppendSinHist "LoadExtraño:" + Textos(J)
            End Select
        End If
    Next J
    
    Terr.Anotar "baj"
    Unload__D = 0
    
    Exit Function
    
ErrLoadPhrase:
    Terr.AppendLog "Unload__D__Err", Terr.ErrToTXT(Err)
End Function


'devuleve matriz de textos incluidos entre llaves
'{no soporte casos de varios inclkuidos dentro de otras {como esta por ejemplo} }
Private Function GetLlaves(fullStr As String) As String()
    
    'si es un archivo lo leo
    If fso.FileExists(fullStr) Then
        Dim TE As TextStream
        Set TE = fso.OpenTextFile(fullStr, ForReading)
            fullStr = TE.ReadAll
        TE.Close
    End If
    
    Dim Ret() As String
    ReDim Ret(0)
    
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
        J = UBound(Ret) + 1
        ReDim Preserve Ret(J)
        Ret(J) = Mid(fullStr, B + 1, B2 - B - 1)
    Loop
    
    GetLlaves = Ret
    
End Function

Public Function GetID1s(i1 As Long, i2 As Long) As String
    GetID1s = ID1s(i1, i2)
End Function

Public Function GetID2s(i1 As Long, i2 As Long) As String
    GetID2s = ID2s(i1, i2)
End Function

Public Function GetComoDev(i1 As Long) As Long
    GetComoDev = ComoDevuelve(i1)
End Function
