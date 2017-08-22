Attribute VB_Name = "Module1"
Public tErr As tbrErrores.clsTbrERR

Public Function Encriptar(Valor, UnEncrypt As Boolean) As String
    'con esta funcion se puede encriptar y desencriptar
    'la uso para el GPF("config")
    
    'para saber si estoy leyendo algo encrytado le pongo algo identificativo
    Dim IdEstaEncryptado As String
    IdEstaEncryptado = "RMLVF"
    'encripta cualquier cosa y la transforma en string
    Dim ToEncrypt As String
    ToEncrypt = CStr(Valor)
    
    Dim Largo As Long, IND As Long, Letra As String, LetraE As String
    Dim FullE As String 'resultado de la encryptacion
    'ver si lo que se ingreso ya esta encrptado
    If UCase(Left(ToEncrypt, Len(IdEstaEncryptado))) = IdEstaEncryptado Then
        'ya esta encriptado
        If UnEncrypt Then
            'DESNCRIPTAR!!!
            'cambiar uno por uno los codigos
            Largo = Len(ToEncrypt)
            'empeiza despues del marcador
            For IND = Len(IdEstaEncryptado) + 1 To Largo
                Letra = Mid(ToEncrypt, IND, 1)
                'pasar todo a una letra distinta. Los saltos de carro no usarlos
                If Asc(Letra) = 0 Then Letra = "0"
                Select Case Letra
                    Case "0"
                        LetraE = vbCrLf
                    Case Else
                        LetraE = Chr(Asc(Letra) - 10)
                End Select
                FullE = FullE + LetraE
            Next
            Encriptar = FullE
        Else
            'no se puede encyprtar lo encryptado
            Encriptar = ToEncrypt
            Exit Function
        End If
    Else
        If UnEncrypt Then
            'no se puede desdencryptar lo desencryptado
            Encriptar = ToEncrypt
            Exit Function
        Else
            'Encriptar!!!!
            'cambiar uno por uno los codigos
            Largo = Len(ToEncrypt)
            For IND = 1 To Largo
                Letra = Mid(ToEncrypt, IND, 1)
                'pasar todo a una letra distinta. Los saltos de carro no usarlos
                Select Case Letra
                    Case vbCrLf ' Or vbCr
                        LetraE = "0"
                    Case Else
                        LetraE = Chr(Asc(Letra) + 10)
                End Select
                FullE = FullE + LetraE
            Next
            Encriptar = IdEstaEncryptado + FullE
        End If
        
    End If
    
End Function

