VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsALotOfPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'dll para almacenar y distribuir imagenes

'arranco pensando que las imagenes esten en otro espacio de memoria que no sea del _
    ejecutable para aliviarlo
'sin saber si el traspaso de uno a otro sea lento

Dim sKey() As String 'identificador unico de cada imagen
'sera su path ya que no se repetira espero

Dim SPic() As stdole.IPictureDisp

'//////////////////////////////////////////////
Dim SPic2() As stdole.StdPicture
Dim SPic3() As IPicture

'vine aqui despues de años (feb 2010) y estaba esto que no permite compilar !?!?
'Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'Public Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long

'//////////////////////////////////////////////
Dim FSO As New Scripting.FileSystemObject


Public Function AddImage(kKey As String, LoadToo As Boolean) As Long
    'si ya esta no lo agrego
    AddImage = 1 'valor si ya estaba
    
    If GetId(kKey) = 0 Then
        Dim H As Long
        H = UBound(sKey) + 1
        ReDim Preserve sKey(H)
        sKey(H) = kKey
        If LoadToo Then
            ReDim Preserve SPic(H)
            'On Local Error GoTo FallaImg
            'Set SPic(H) =  New stdole.StdPicture
            Set SPic(H) = LoadPicture(sKey(H))
        End If
        AddImage = 0
    End If
    Exit Function
    
FallaImg:
    'eliminarla!
    H = UBound(sKey) - 1
    ReDim Preserve sKey(H)
    ReDim Preserve SPic(H)
    AddImage = -1
End Function

Public Function GetPicture(kKey As String, kKeyIfNoExists As String) As StdPicture
    
    Dim H As Long, IACT As String 'imagen actual buscada
    
    IACT = kKey
PROBAR:
    'si pasa un numero como parametro lo uso como indice!
    If IsNumeric(IACT) Then
        H = CLng(IACT)
    Else
        H = GetId(IACT)
    End If
    
    If (H > 0) And (H <= UBound(SPic)) Then
        Set GetPicture = SPic(H)
        Exit Function
    Else
        'puede que no este en la lista de imagenes por que trabaja a pedido
        If FSO.FileExists(IACT) Then
            On Local Error GoTo ImgNoValid
            'ver si es buena!!!!
            Set GetPicture = LoadPicture(IACT)
            Exit Function
        End If
    End If
    
    'si llego hasta aqui ni esta en la base de imagenes ni el archivo existe
    'le doy una oportunidad a la imagen 2 y si no salgo
    If IACT = kKeyIfNoExists Then
        Set GetPicture = LoadPicture()
        Exit Function
    Else
        IACT = kKeyIfNoExists
        GoTo PROBAR
    End If
    
    Exit Function
    
ImgNoValid:
    If IACT = kKey Then GoTo PROBAR

End Function

Private Sub Class_Initialize()
    ClearAll
End Sub

Public Function ClearAll()
    ReDim sKey(0)
    ReDim SPic(0)
End Function

Private Function GetId(kKey As String)
    'ver si existe
    GetId = 0 'bandera de que no existe
    For H = 1 To UBound(sKey)
        If LCase(sKey(H)) = LCase(kKey) Then
            GetId = H 'ya existe
            Exit Function
        End If
    Next H
End Function
