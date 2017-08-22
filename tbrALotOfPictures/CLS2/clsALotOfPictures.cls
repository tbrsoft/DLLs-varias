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
'// ------------------------------------------------------------------------------------- //
'// Modulo para cargar Mapas de Bits en memoria y funcion de lectura de pixeles para usar //
'// como Mapas de Durezas.                                                                //
'//                                                                                       //
'// [EX3] - José Miguel Sánchez Fernández                                                 //
'// ------------------------------------------------------------------------------------- //

Option Explicit

'// Declaraciones del API de Windows //

'// CreateCompatibleDC() - Crea un Contexto de Dispositivo (DC) compatible con el dispositivo especificado:
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

' // DeleteDC() - Borra el Contexto de Dispositivo especificado:
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'// SelectObject() - Selecciona un objeto dentro del contexto de dispositivo especificado. El nuevo objeto sustituye el objeto anterior del mismo tipo:
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'// DeleteObject() - Libera todos los recursos del sistema asociados al objeto. Despues el objeto es eliminado, el identificador especificado deja de ser valido:
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'// GDIGetPixel() alias GetPixel() - Devuelve los canales RGB del color del pixel especificado por las coordenadas:
Private Declare Function GDIGetPixel Lib "gdi32" Alias "GetPixel" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

'// Estructura que representa un mapa de bits en memoria:
Private Type BMP_Handle
    sPath As String
    hDC As Long '// Idenficicador (handle) del Contexto de Dispositivo (hDC)
    Surface As Picture '// Superficie definida por el objeto Picture para almacenar los mapas de bits.
End Type

'// Array que almacenara los mapas de bits:
Private MyBitMap() As BMP_Handle

'// Carga un archivo de imagen BMP, DIB, GIF, JPEG, WMF mediante el objeto _
    Picture de la libreria del Visual Basic 6.0: (Antencion: Objeto Picture, no confundir con Control PictureBox)



'***************************************************************************
'dll para almacenar y distribuir imagenes

'arranco pensando que las imagenes esten en otro espacio de memoria que no sea del _
    ejecutable para aliviarlo
'sin saber si el traspaso de uno a otro sea lento

Dim sKey() As String 'identificador unico de cada imagen
'sera su path ya que no se repetira espero

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
            On Local Error GoTo FallaImg
            Load_BitMap kKey
        End If
        AddImage = 0
    End If
    Exit Function
    
FallaImg:
    'eliminarla!
    H = UBound(sKey) - 1
    ReDim Preserve sKey(H)
    AddImage = -1
End Function

Public Function GetPicture(kKey As String, kKeyIfNoExists As String) As StdPicture
    Dim H As Long
    H = GetId(kKey)
    If H > 0 Then 'si existe en la lista ...
        'ver si estaba cargado o no!
        If (UBound(MyBitMap) > 0) Then
            Set GetPicture = SetPicture(GetId(kKey))
        Else
            'si no esta no quiere decir que no exista la imagene
            'puede ser que se cargue en demanda!
            If FSO.FileExists(kKey) Then
                'carga a pedido!
                Set GetPicture = LoadPicture(kKey)
            Else
                Set GetPicture = LoadPicture()
            End If
        End If
    Else
'        H = GetId(kKeyIfNoExists)
'        If H > 0 Then
'            Set GetPicture = SPic(GetId(kKey))
'        Else
'            Set GetPicture = LoadPicture 'NO HAY NI LA ORIGINAL NI LA PREDETERMINADA!!
'        End If

    End If
End Function

Private Sub Class_Initialize()
    ClearAll
End Sub

Public Function ClearAll()
    ReDim sKey(0)
    'ReDim SPic(0)
    ReDim MyBitMap(0)
End Function

Private Function GetId(kKey As String)
    'ver si existe
    GetId = 0 'bandera de que no existe
    Dim H As Long
    For H = 1 To UBound(sKey)
        If LCase(sKey(H)) = LCase(kKey) Then
            GetId = H 'ya existe
            Exit Function
        End If
    Next H
End Function

Public Function SetPicture(sPathF As String) As StdPicture
    Set SetPicture = MyBitMap(getI(sPathF)).Surface
End Function

Private Function getI(sFile As String)
    Dim H As Long
    For H = 1 To UBound(MyBitMap)
        If LCase(sFile) = LCase(MyBitMap(H).sPath) Then
            getI = H
            Exit Function
        End If
    Next H
    getI = 0
End Function

Public Function Load_BitMap(FileName As String) As Long
    On Error GoTo ErrOut

    Dim i As Long '// Ubicacion del mapa de bits en el array.
    
    i = Get_FreeHandle '// Obtenemos una ubicacion valida (o vacia si la hubiese) en el array para almacenar el mapa de bits.
    
    With MyBitMap(i)
        .sPath = FileName
        .hDC = CreateCompatibleDC(0) '// Creamos un DC compatible con el dispositivo por defecto.
        Set .Surface = LoadPicture(FileName) '// Cargamos la imagen en la superficie mediante la funcion LoadPicture() de la libreria de Visual Basic 6.0
    
        Call SelectObject(.hDC, .Surface.Handle) '// Creamos un nuevo objeto con la informacion del mapa de bits.
    
        Load_BitMap = i '// Devolvemos la ubicacion del mapa de bits en el array a modo de indentificador con el que se accedera despues para leer los pixeles y para borrar la imagen de la memoria.
    
    End With
    
ErrOut:
    
End Function

Public Sub UnloadAll()
    Dim H As Long
    For H = 1 To UBound(MyBitMap)
        Unload_BitMap H
    Next H
End Sub

'// Descarga un mapa de bits de la memoria:
Public Sub Unload_BitMap(Handle As Long)
On Error GoTo ErrOut

    '// Si el identificador corresponde con el ultimo elemento del array...
    If Handle = UBound(MyBitMap) Then
        If Handle = 0 Then '//Si el identificador es el primer elemento del array...
            Erase MyBitMap '// Eliminamos por completo el array liberando la memoria ocupada.

        Else '// Si se trata del ultimo elemento del array entonces redimensionamos el array con un elemento menos liberando la memoria ocupada por el:
            ReDim Preserve MyBitMap(UBound(MyBitMap) - 1) As BMP_Handle

        End If

    Else '// Si se trata de un elemento intermedio del array liberamos su memoria pero conservamos su espacio para ser usado en la proxima carga de Load_BitMap():

        With MyBitMap(Handle)
            Call DeleteObject(.Surface.Handle) '// Eliminamos el objeto liberando su memoria.
            Call DeleteDC(.hDC) '// Eliminamos el Contexto de Dispostivo.
            Set .Surface = Nothing '// Liberamos la superficie.

        End With


    End If

ErrOut:

End Sub

'// Lee el pixel de un mapa de bits cargado en memoria:
Public Function GetPixel(Handle As Long, X As Long, Y As Long) As Long
    GetPixel = GDIGetPixel(MyBitMap(Handle).hDC, X, Y)
End Function

'// Devuelve una ubicacion nueva o vacia en el array para almacenar un nuevo mapa de bits en memoria:
Private Function Get_FreeHandle() As Long
On Error Resume Next

Dim i As Long, j As Long

i = UBound(MyBitMap) '// Obtenemos la cantidad de elemenos del array.

If Err.Number = 9 Then '// Si el array no contiene elementos...
    ReDim MyBitMap(0) As BMP_Handle '// Redimensionamos el array a 0 creando un elemento vacio.

    Get_FreeHandle = 0 '// Devolvemos la posicion del elemento que seria 0 al ser el primero.

Else '// Si el array contenia elementos buscamos el primer elemento vacio que exista y se recicla:
    For j = 0 To i
        If MyBitMap(j).Surface Is Nothing Then '// Si la superficie esta liberada hemos encontrado un elemento vacio.
            Get_FreeHandle = j '// Devolvemos la posicion del elemento.
            Exit For '// Salimos del bucle.

        End If

    Next j

    '// Si no existia ningun elemento vacio se añade uno nuevo al array:
    ReDim Preserve MyBitMap(i + 1) As BMP_Handle '// Redimensionamos el array incrementando su indice:
    Get_FreeHandle = i + 1 '// Devolvemos la posicion del nuevo elemento.

End If

End Function

