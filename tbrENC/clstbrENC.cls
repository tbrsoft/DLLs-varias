VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clstbrENC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub Encriptar(Invertido As Boolean, Clave As String, ArchivoIn As String, ArchivoOut As String, Optional Prefix As String = "")
'Cargo los datos
Dim Buffer() As Byte
'Dim NuevoDato() As Byte

Dim f As Integer

Dim xClave() As Byte
ReDim xClave(Len(Clave))
For f = 1 To Len(Clave)
    xClave(f) = Asc(Mid(Clave, f, 1))
Next f


'Dato = Space(FileLen(ArchivoIn))

ReDim Buffer(FileLen(ArchivoIn))
Open ArchivoIn For Binary As #1
    If Invertido = True Then
        Get #1, Len(Prefix) + 1, Buffer
    Else
        Get #1, 1, Buffer
    End If
Close

'Encripto

Dim Char1 As Integer 'Caracter Original
Dim Char2 As Integer 'Caracter ya Modificado (char1+char3) o (char1-char3)
Dim Char3 As Integer 'Caracter de la Clave

'Voy dando vueltas por la clave asi que necesito un indice
Dim ContadorClave As Integer 'Indice de la clave
ContadorClave = 1

Dim i As Long
i = 0

'For i = 0 To UBound(Buffer)
Do
Repe:
    If i > UBound(Buffer) Then Exit Do
    
    'Form1.Caption = i
    '----Para la barra del Formulario----
    'Dim PorC As Integer
    'Dim BW As Integer
    'PorC = i / UBound(Buffer)
    'BW = PorC * Form1.pBar.Width
    'PorC = PorC * 100
    'Form1.pBar.Line (0, 0)-(BW, 20), vbBlue, BF
    'Form1.pBar.Cls
    'Form1.pBar.Print CStr(PorC) + "%"
    'Form1.pBar.Refresh
    
    'Form1.Caption = CStr(PorC) + "%"
    '------------------------------------

    Char1 = Buffer(i)

    Char3 = xClave(ContadorClave)

    If Invertido = True Then
        Char2 = Char1 - Char3
    Else
        Char2 = Char1 + Char3
    End If


    If Char2 < 0 Then
        Char2 = 256 + Char2
    End If

    If Char2 > 255 Then
        Char2 = Char2 Mod 256
    End If



    'NuevoDato(i) = Char2
    Buffer(i) = Char2
    'DoEvents


    ContadorClave = ContadorClave + 1
    If ContadorClave > UBound(xClave) Then
        'i = i + ContadorClave
        ContadorClave = 1
        GoTo Repe
    End If


'Next i
    i = i + 1
Loop
'Form1.pBar.Cls


'Grabo el archivo
If Dir(ArchivoOut) <> "" Then Kill ArchivoOut
Open ArchivoOut For Binary As #1
    'Put #1, 1, NuevoDato
    If Invertido = True Then
        Put #1, 1, Buffer
    Else
        Put #1, 1, Prefix
        Put #1, Len(Prefix) + 1, Buffer
    End If
Close

End Sub
'[ + ]

Public Sub EncriptarTXT(Invertido As Boolean, Clave As String, ByRef Buffer() As Byte)
'Cargo los datos

Dim f As Integer

Dim xClave() As Byte
ReDim xClave(Len(Clave))
For f = 1 To Len(Clave)
    xClave(f) = Asc(Mid(Clave, f, 1))
Next f


'Encripto

Dim Char1 As Integer 'Caracter Original
Dim Char2 As Integer 'Caracter ya Modificado (char1+char3) o (char1-char3)
Dim Char3 As Integer 'Caracter de la Clave

'Voy dando vueltas por la clave asi que necesito un indice
Dim ContadorClave As Integer 'Indice de la clave
ContadorClave = 1

Dim i As Long
i = 0

'For i = 0 To UBound(Buffer)
Do
R2epe:
    
    If i > UBound(Buffer) Then Exit Do
    

    Char1 = Buffer(i)

    Char3 = xClave(ContadorClave)

    If Invertido = True Then
        Char2 = Char1 - Char3
    Else
        Char2 = Char1 + Char3
    End If


    If Char2 < 0 Then
        Char2 = 256 + Char2
    End If

    If Char2 > 255 Then
        Char2 = Char2 Mod 256
    End If



    'NuevoDato(i) = Char2
    Buffer(i) = Char2
    'DoEvents


    ContadorClave = ContadorClave + 1
    If ContadorClave > UBound(xClave) Then
        'i = i + ContadorClave
        ContadorClave = 1
        GoTo R2epe
    End If


'Next i
    i = i + 1
Loop


End Sub


