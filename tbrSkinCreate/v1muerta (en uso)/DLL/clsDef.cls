VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clase que se carga y graba definicioes
'se refiere a la lista de los objetos

Public Enum sTRANS 'propiedades de la transparecia de la imagen
    Obligatoria = 1
    Opcional = 2
    Prohibida = 3
End Enum

Public Type defIMG
    sPath As String
    sName As String
    sMinWidth As Long
    sMaxWidth As Long
    sMinHeight As Long
    sMaxHeight As Long
    sAnchoFinal As Long
    sAltoFinal As Long
    sCoef As Single    'coeficiente acho/alto (o alreves)
    sTransparencia As sTRANS
    sTrasnparenciaDescripcion As String
    sMinMargenDerechoTrans As Long
    sMaxMargenDerechoTrans As Long
    sMargenDerechoTransFINAL As Long
    sMinMargenIzquierdoTrans As Long
    sMaxMargenIzquierdoTrans As Long
    sMargenIzquierdoTransFINAL As Long
    sMinMargenSuperiorTrans As Long
    sMaxMargenSuperiorTrans As Long
    sMargenSuperiorTransFINAL As Long
    sMinMargenInferiorTrans As Long
    sMaxMargenInferiorTrans As Long
    sMargenInferiorTransFINAL As Long
End Type

Public Type defColors
    sName As String
    sColor As Long
End Type

Private dImgs() As defIMG
Private dColors() As defColors

Public Function getColor(byName As String) As Long
    'devuelve un path para que use una imagen empaquetada!
    'pasar por todas las imagenes hasta encontrar coincidencia!
    Dim J As Long, Encontro As Long
    Encontro = 0
    For J = 1 To UBound(dColors)
        If LCase(byName) = LCase(dColors(J).sName) Then
            Encontro = dColors(J).sColor
            Exit For
        End If
    Next J
    getColor = Encontro
End Function

Public Function getColorById(i As Long) As Long
    If i > UBound(dColors) Then Exit Function
    getColorById = dColors(i).sColor
End Function

Public Function getNameColor(idC As Long) As String
    If idC > UBound(dColors) Then Exit Function
    getNameColor = dColors(idC).sName
End Function

Public Function GetIndexImage(byName As String) As Long
    'devuelve el indice en las matrices de la imagen solicitada
    '-1 no encontro!
    Dim J As Long
    GetIndexImage = -1
    For J = 1 To UBound(dImgs)
        If LCase(byName) = LCase(dImgs(J).sName) Then
            GetIndexImage = J
            Exit For
        End If
    Next J
End Function

Public Function getImagePath(byName As String) As String
    'devuelve un path para que use una imagen empaquetada!
    'pasar por todas las imagenes hasta encontrar coincidencia!
    Dim J As Long, Encontro As String
    Encontro = ""
    For J = 1 To UBound(dImgs)
        If LCase(byName) = LCase(dImgs(J).sName) Then
            Encontro = dImgs(J).sPath  'al abrir ya se descomprimio y actualizo pathImages !!!
            Exit For
        End If
    Next J
    getImagePath = Encontro
End Function
    
Public Sub Clean()
    CantImgs = 0
    ReDim Preserve dImgs(0)
    ReDim Preserve dColors(0)
End Sub

Public Sub RemoveImage(i As Long)
    If i > UBound(dImgs) Then Exit Sub
    If UBound(dImgs) = 0 Then Exit Sub
    
    Dim J As Long
    For J = i To UBound(dImgs) - 1
        dImgs(i) = dImgs(i + 1)
    Next J
    
    J = UBound(dImgs) - 1
    
    ReDim Preserve dImgs(J)
    
End Sub

Public Sub RemoveColor(i As Long)
    If i > UBound(dColors) Then Exit Sub
    If UBound(dColors) = 0 Then Exit Sub
    
    'directamente lo reemplazo con el de arriba
    Dim J As Long
    For J = i To UBound(dColors) - 1
        dColors(i) = dColors(i + 1)
    Next J
    
    J = UBound(dColors) - 1
    
    ReDim Preserve dColors(J)
End Sub

Public Function GetIndexColor(idColor As String) As Long
    'devuelve el indice en las matrices de la imagen solicitada
    '-1 no encontro!
    Dim J As Long
    GetIndexColor = -1
    For J = 1 To UBound(dColors)
        If LCase(idColor) = LCase(dColors(J).sName) Then
            GetIndexColor = J
            Exit For
        End If
    Next J
End Function

Public Function GetCantImgs()
    GetCantImgs = UBound(dImgs)
End Function

Public Function GetCantColores()
    GetCantColores = UBound(dColors)
End Function

'**********************************************
Public Sub ChgPathImage(i As Long, valU As String)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sPath = valU
End Sub

Public Sub AddImage(sPath2 As String)
    Dim J As Long
    J = UBound(dImgs) + 1
    ReDim Preserve dImgs(J)
    dImgs(J).sPath = sPath2
End Sub

Public Sub AddImageFull(defImage As defIMG)
    Dim J As Long
    J = UBound(dImgs) + 1
    ReDim Preserve dImgs(J)
    dImgs(J) = defImage
End Sub

Public Function getDefImage(i As Long) As defIMG
    If i > UBound(dImgs) Then Exit Function
    If i = -1 Then i = UBound(dImgs)
    getDefImage = dImgs(i)
End Function

Public Sub AddColor(sColor2 As Long, sName2 As String)
    Dim J As Long
    J = UBound(dColors) + 1
    ReDim Preserve dColors(J)
    dColors(J).sColor = sColor2
    dColors(J).sName = sName2
End Sub

Public Sub AddColorOnlyName(sName2 As String)
    Dim J As Long
    J = UBound(dColors) + 1
    ReDim Preserve dColors(J)
    dColors(J).sName = sName2
End Sub


Public Sub DefineMinWidth(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMinWidth = valU
End Sub

Public Sub DefineMaxWidth(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMaxWidth = valU
End Sub

Public Sub DefineMinHeight(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMinHeight = valU
End Sub

Public Sub DefineMaxHeight(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMaxHeight = valU
End Sub

Public Sub DefineCoef(i As Long, valU As Single)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sCoef = valU
End Sub

Public Sub DefineTranspType(i As Long, valU As sTRANS)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sTransparencia = valU
End Sub

Public Sub DefineNameImage(i As Long, valU As String)
    If i > UBound(dImgs) Then Exit Sub
    '-1 es en el ultimo (cuando estoy cargando viene joia
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sName = valU
End Sub

Public Sub DefineTranspDescripcion(i As Long, valU As String)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sTrasnparenciaDescripcion = valU
End Sub

Public Sub DefineMinMargenDerechoTrans(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMinMargenDerechoTrans = valU
End Sub

Public Sub DefineMinMargenIzquierdoTrans(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMinMargenIzquierdoTrans = valU
End Sub

Public Sub DefineMinMargenSuperiorTrans(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMinMargenSuperiorTrans = valU
End Sub

Public Sub DefineMinMargenInferiorTrans(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMinMargenInferiorTrans = valU
End Sub

Public Sub DefineMaxMargenDerechoTrans(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMaxMargenDerechoTrans = valU
End Sub

Public Sub DefineMaxMargenIzquierdoTrans(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMaxMargenIzquierdoTrans = valU
End Sub

Public Sub DefineMaxMargenSuperiorTrans(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMaxMargenSuperiorTrans = valU
End Sub

Public Sub DefineMaxMargenInferiorTrans(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMaxMargenInferiorTrans = valU
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub DefineColorValue(i As Long, valU As Long)
    If i > UBound(dColors) Then Exit Sub
    If i = -1 Then i = UBound(dColors)
    dColors(i).sColor = valU
End Sub

Public Sub DefineColorName(i As Long, newName As String)
    If i > UBound(dColors) Then Exit Sub
    If i = -1 Then i = UBound(dColors)
    dColors(i).sName = newName
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function GetpathImage(i As Long) As String
    If i > UBound(dImgs) Then Exit Function
    GetpathImage = dImgs(i).sPath
End Function

Public Function GetMinWidth(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMinWidth = dImgs(i).sMinWidth
End Function

Public Function GetMaxWidth(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMaxWidth = dImgs(i).sMaxWidth
End Function

Public Function GetMinHeight(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMinHeight = dImgs(i).sMinHeight
End Function

Public Function GetMaxHeight(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMaxHeight = dImgs(i).sMaxHeight
End Function

Public Function GetCoef(i As Long) As Single
    If i > UBound(dImgs) Then Exit Function
    GetCoef = dImgs(i).sCoef
End Function

Public Function GetTranspType(i As Long) As sTRANS
    If i > UBound(dImgs) Then Exit Function
    If dImgs(i).sTransparencia = 0 Then
        GetTranspType = Opcional
    Else
        GetTranspType = dImgs(i).sTransparencia
    End If
End Function

Public Function GetNameImage(i As Long) As String
    If i > UBound(dImgs) Then Exit Function
    GetNameImage = dImgs(i).sName
End Function

Public Function GetTranspDescripcion(i As Long) As String
    If i > UBound(dImgs) Then Exit Function
    GetTranspDescripcion = dImgs(i).sTrasnparenciaDescripcion
End Function

Public Function GetMinMargenDerechoTrans(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMinMargenDerechoTrans = dImgs(i).sMinMargenDerechoTrans
End Function

Public Function GetMinMargenIzquierdoTrans(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMinMargenIzquierdoTrans = dImgs(i).sMinMargenIzquierdoTrans
End Function

Public Function GetMinMargenSuperiorTrans(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMinMargenSuperiorTrans = dImgs(i).sMinMargenSuperiorTrans
End Function

Public Function GetMinMargenInferiorTrans(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMinMargenInferiorTrans = dImgs(i).sMinMargenInferiorTrans
End Function

Public Function GetMaxMargenDerechoTrans(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMaxMargenDerechoTrans = dImgs(i).sMaxMargenDerechoTrans
End Function

Public Function GetMaxMargenIzquierdoTrans(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMaxMargenIzquierdoTrans = dImgs(i).sMaxMargenIzquierdoTrans
End Function

Public Function GetMaxMargenSuperiorTrans(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMaxMargenSuperiorTrans = dImgs(i).sMaxMargenSuperiorTrans
End Function

Public Function GetMaxMargenInferiorTrans(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetMaxMargenInferiorTrans = dImgs(i).sMaxMargenInferiorTrans
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub DefineFinalHeight(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sAltoFinal = valU
End Sub

Public Sub DefineFinalWidth(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sAnchoFinal = valU
End Sub

Public Sub DefineFinalMargenDerechoTra(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMargenDerechoTransFINAL = valU
End Sub

Public Sub DefineFinalMargenIzquierdoTra(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMargenIzquierdoTransFINAL = valU
End Sub

Public Sub DefineFinalMargenSuperiorTra(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMargenSuperiorTransFINAL = valU
End Sub

Public Sub DefineFinalMargenInferiorTra(i As Long, valU As Long)
    If i > UBound(dImgs) Then Exit Sub
    If i = -1 Then i = UBound(dImgs)
    dImgs(i).sMargenInferiorTransFINAL = valU
End Sub

'******************+++++++++*************************

Public Function GetFinalHeight(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetFinalHeight = dImgs(i).sAltoFinal
End Function

Public Function GetFinalWidth(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetFinalWidth = dImgs(i).sAnchoFinal
End Function

Public Function GetFinalMargenDerechoTra(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetFinalMargenDerechoTra = dImgs(i).sMargenDerechoTransFINAL
End Function

Public Function GetFinalMargenIzquierdoTra(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetFinalMargenIzquierdoTra = dImgs(i).sMargenIzquierdoTransFINAL
End Function

Public Function GetFinalMargenSuperiorTra(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetFinalMargenSuperiorTra = dImgs(i).sMargenSuperiorTransFINAL
End Function

Public Function GetFinalMargenInferiorTra(i As Long) As Long
    If i > UBound(dImgs) Then Exit Function
    GetFinalMargenInferiorTra = dImgs(i).sMargenInferiorTransFINAL
End Function

Public Function toString(i As Long) As String

    If i > UBound(dImgs) Then Exit Function
    If i = -1 Then i = UBound(dImgs)
    
    s = "**********************************" + vbCrLf + _
        "Name: " + dImgs(i).sName + vbCrLf + _
        "**********************************" + vbCrLf + _
        "Path: " + dImgs(i).sPath + vbCrLf + _
        "Wi: " + CStr(dImgs(i).sMinWidth) + "-" + CStr(dImgs(i).sMaxWidth) + vbCrLf + _
        "He: " + CStr(dImgs(i).sMinHeight) + "-" + CStr(dImgs(i).sMaxHeight) + vbCrLf + _
        "Ancho-Alto: " + CStr(dImgs(i).sAnchoFinal) + "-" + CStr(dImgs(i).sAltoFinal) + vbCrLf + _
        "Coef: " + CStr(dImgs(i).sCoef) + vbCrLf + _
        "Tansp: (" + CStr(dImgs(i).sTransparencia) + ")" + dImgs(i).sTrasnparenciaDescripcion + vbCrLf + _
        "MargenDER: " + CStr(dImgs(i).sMinMargenDerechoTrans) + "-" + CStr(dImgs(i).sMaxMargenDerechoTrans) + "-" + CStr(dImgs(i).sMargenDerechoTransFINAL) + vbCrLf + _
        "MargenIZQ: " + CStr(dImgs(i).sMinMargenIzquierdoTrans) + "-" + CStr(dImgs(i).sMaxMargenIzquierdoTrans) + "-" + CStr(dImgs(i).sMargenIzquierdoTransFINAL) + vbCrLf + _
        "MargenSUP: " + CStr(dImgs(i).sMinMargenSuperiorTrans) + "-" + CStr(dImgs(i).sMaxMargenSuperiorTrans) + "-" + CStr(dImgs(i).sMargenSuperiorTransFINAL) + vbCrLf + _
        "MargenINF: " + CStr(dImgs(i).sMinMargenInferiorTrans) + "-" + CStr(dImgs(i).sMaxMargenInferiorTrans) + "-" + CStr(dImgs(i).sMargenInferiorTransFINAL)
    
    toString = s
    
End Function
