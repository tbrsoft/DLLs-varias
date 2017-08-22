VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Translator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'hay que agregar la referencia a Microsoft Scripting Runtime

Dim vLanguage As String
Dim vDictionary As New Dictionary

Public Property Get Dictionary() As Dictionary
    Set Dictionary = vDictionary
End Property

Public Property Get Language() As String
    'ver si hay q devolver default en vez de cast.
    Language = IIf(vLanguage <> "", vLanguage, "Castellano")
End Property

Public Property Let Language(pValor As String)
    vLanguage = pValor
    LoadDictionary pValor
    
End Property

Public Property Get Languages() As String()
    Dim fso
    Dim folder
    Dim aux As String
    
    Set fso = CreateObject("scripting.filesystemobject")
    Set folder = fso.GetFolder(App.path + "\idiomas")
    Dim f
    For Each f In folder.Files
        aux = aux + Left(f.Name, Len(f.Name) - 4) + "*"
    Next
    Languages = Split(Left(aux, Len(aux) - 1), "*")
End Property

Public Function Translate(pSpanishWord As String, _
    Optional sVARS As String = "") As String
    
    'busca en el diccionario el texto el texto y devuelve la _
        traduccion segun el diccionario cargado
    
    On Error GoTo e
    Dim aux As String
    aux = vDictionary.Item(pSpanishWord)
    If aux <> "" Then
        Translate = aux
    Else
        'si no encuentra la palabra, devuelve la misma
        Translate = pSpanishWord
    End If
    Exit Function
e:
'si no encuentra la palabra, devuelve la misma
    Translate = pSpanishWord
    
End Function

Private Sub LoadDictionary(pLanguage As String)
    On Error GoTo e
   
    Dim s As String
    Dim aux() As String
    Dim par() As String
    
    s = LeerArchivo(App.path + "\idiomas\" + pLanguage + ".lan")
    'los archivos tiene lineas separadas por chr(30) y dentro de _
        cada linea separada con un chr(31) entre el texto original y el traducido
    aux = Split(s, Chr$(30))
    vDictionary.RemoveAll
    For I = 0 To UBound(aux)
        If aux(I) <> "" Then
            par = Split(aux(I), Chr$(31))
            vDictionary.Add par(0), par(1)
        End If
    Next
        
    Exit Sub
e:
    MsgBox "No se encontro el archivo de idioma."
End Sub

Private Function LeerArchivo(path As String) As String
    Dim fso 'As FileSystemObject
    Dim f
        
    Dim s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If path = "" Then
        LeerArchivo = ""
    Else
        Set f = fso.GetFile(path)
    End If
    Set ts = f.OpenAsTextStream(1)
    
    s = ts.ReadAll
    
    ts.Close
    LeerArchivo = s
End Function

