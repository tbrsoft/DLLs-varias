VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClaveTbr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function Encriptarr(Texto As String, Clave As String, _
                        Encriptando As Boolean) As String
                        
    'para que no haya overflow le pongo
    Dim ValorDeNumNomb As Double, ValorCadena As Double, X As Long
    
    'sumar el valor nuemrioco de la suma del codigo de cada letra
    'siempre sumandole 13 (imagino que para despistar)
    For X = 1 To Len(Clave)
        ValorDeNumNomb = ValorDeNumNomb + Asc(Mid$(Clave, X, 1)) + 13
    Next X
    
    For X = 1 To Len(Texto)
        'encripta o desencripta el strOriginal
        If Encriptando = True Then
            ValorCadena = Asc(Mid$(Texto, X, 1)) + ValorDeNumNomb + (X ^ 2)
            Do While ValorCadena > 255
                ValorCadena = ValorCadena - 255
            Loop
            
        Else
            ValorCadena = Asc(Mid$(Texto, X, 1)) - _
                ValorDeNumNomb - (X ^ 2)
                
            Do While ValorCadena <= 0
                ValorCadena = ValorCadena + 255
            Loop
            
        End If
     
        Encriptarr = Encriptarr & Chr(ValorCadena)
    Next X

End Function


