VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPakageSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sNamePackage As String
Private sNameSKIN As String
Private sPathImages() As String
Private sNameImages() As String

'para definir y controlar
Private sMinWidth() As Long: Private sMaxWidth() As Long: Private sMinHeight() As Long: Private sMaxHeight() As Long
'solo lo usa el skin
Private sAnchoFinal() As Long: Private sAltoFinal() As Long

Private sCoef() As Single

Public Enum sTRANS
    Obligatoria = 1
    Opcional = 2
    Prohibida = 3
End Enum

Private sTransparencia() As sTRANS

Private sTrasnparenciaDescripcion() As String

Private sMinMargenDerechoTrans() As Long: Private sMaxMargenDerechoTrans() As Long
Private sMargenDerechoTransFINAL() As Long
Private sMinMargenIzquierdoTrans() As Long: Private sMaxMargenIzquierdoTrans() As Long
Private sMargenIzquierdoTransFINAL() As Long
Private sMinMargenSuperiorTrans() As Long: Private sMaxMargenSuperiorTrans() As Long
Private sMargenSuperiorTransFINAL() As Long
Private sMinMargenInferiorTrans() As Long: Private sMaxMargenInferiorTrans() As Long
Private sMargenInferiorTransFINAL() As Long

Private CantImgs As Long 'contador de imagenes

Public Function GetCantImgs()
    GetCantImgs = CantImgs
End Function

Public Sub AbrirSKI_(ArchSKI_ As String)
    'descomprimir todo
    'meter el .DEF en el modulo
    'indicar el path de lo descomprimido
    Dim FSO As New Scripting.FileSystemObject
    Dim JS As New tbrJUSE.clsJUSE
    'leerlo
    JS.ReadFile ArchSKI_
    'extraer todo
    Dim ET As String 'extraer todo en ...
    ET = FSO.GetParentFolderName(ArchSKI_) + "\" + FSO.GetBaseName(ArchSKI_) + "\"
    
    If FSO.FolderExists(ET) = False Then FSO.CreateFolder ET
    Dim A As Long
    For A = 1 To JS.CantArchs
        JS.Extract ET, A
    Next A
    'cerrar todo
    Set JS = Nothing
    
    loadDef ET + FSO.GetBaseName(ArchSKI_) + ".def"
    
    Set FSO = Nothing
End Sub

Public Sub AbrirSKIN(ArchSKIN As String)
    'descomprimir todo
    'meter el .DEF en el modulo
    'indicar el path de lo descomprimido
    Dim FSO As New Scripting.FileSystemObject
    Dim JS As New tbrJUSE.clsJUSE
    'leerlo
    JS.ReadFile ArchSKIN
    'extraer todo
    Dim ET As String 'extraer todo en ...
    ET = FSO.GetParentFolderName(ArchSKIN) + "\" + FSO.GetBaseName(ArchSKIN) + "\"
    
    If FSO.FolderExists(ET) = False Then FSO.CreateFolder ET
    Dim A As Long
    For A = 1 To JS.CantArchs
        JS.Extract ET, A
    Next A
    'cerrar todo
    Set JS = Nothing
    
    loadDef ET + FSO.GetBaseName(ArchSKIN) + ".defN"
    
    Set FSO = Nothing
End Sub

Public Sub loadDef(sArch As String)
    
    'abre validando que todo exista!
    TodoACero 'limpio todas las matrices
    
    'monto la definicion primero
    Dim FSO As New Scripting.FileSystemObject
    
    'esta todo dentro de la carpeta donde esta el def
    Dim FOL As String
    
    FOL = FSO.GetParentFolderName(sArch)
    If Right(FOL, 1) <> "\" Then FOL = FOL + "\"
    
    Dim TE As TextStream, tmpT As String
    Set TE = FSO.OpenTextFile(sArch)
        tmpT = TE.ReadAll
    TE.Close
    
    Dim Lineas() As String
    Lineas = Split(tmpT, Chr(5))
    Dim J As Long
    Dim Tit() As String, Resto As String 'titulo y resto de cada linea
    For J = 0 To UBound(Lineas)
        If Lineas(J) = "" Then GoTo SIGLINEA
        Tit = Split(Lineas(J), ":")
        Resto = Mid(Lineas(J), Len(Tit(0)) + 2, Len(Lineas(J)) - Len(Tit(0)) - 1) '+2 es para que saque los ":"
        'ej: minWidth:334
            'tit=minwidth
            'len(tit)=8 (+2=10)
            'len linea(j)=12
            'mid start=10 len=3 'JOIA
            
        Select Case LCase(Tit(0))
            'titulo del archivo, va al testo arriba
            Case "package": sNamePackage = Resto
            'empieza una imagen, agregarla para que se vayan creando los indices!
            Case "imagepath": AddImage FOL + FSO.GetBaseName(Resto) + "." + FSO.GetExtensionName(Resto)
            'id unico de la imagen
            Case "image": DefineNameImage -1, Resto
            Case "minwidth": DefineMinWidth -1, CLng(Resto)
            Case "maxwidth": DefineMaxWidth -1, CLng(Resto)
            
            Case "minheight": DefineMinHeight -1, CLng(Resto)
            Case "maxheight": DefineMaxHeight -1, CLng(Resto)
            Case "coef": DefineCoef -1, CSng(Resto)
            Case "transp": DefineTranspType -1, CLng(Resto)
            Case "transpdet": DefineTranspDescripcion -1, Resto
            
            Case LCase("MinMargenDerechoTrans"): DefineMinMargenDerechoTrans -1, CLng(Resto)
            Case LCase("MinMargenIzquierdoTrans"): DefineMinMargenIzquierdoTrans -1, CLng(Resto)
            Case LCase("MinMargenSuperiorTrans"): DefineMinMargenSuperiorTrans -1, CLng(Resto)
            Case LCase("MinMargenInferiorTrans"): DefineMinMargenInferiorTrans -1, CLng(Resto)
            
            Case LCase("MaxMargenDerechoTrans"): DefineMaxMargenDerechoTrans -1, CLng(Resto)
            Case LCase("MaxMargenIzquierdoTrans"): DefineMaxMargenIzquierdoTrans -1, CLng(Resto)
            Case LCase("MaxMargenSuperiorTrans"): DefineMaxMargenSuperiorTrans -1, CLng(Resto)
            Case LCase("MaxMargenInferiorTrans"): DefineMaxMargenInferiorTrans -1, CLng(Resto)
        
            'solo para skins!
            Case LCase("FINALHEIGHT"): DefineFinalHeight -1, CLng(Resto)
            Case LCase("FINALWIDTH"): DefineFinalWidth -1, CLng(Resto)
            Case LCase("FINALMargenDerechoTRA"): DefineFinalMargenDerechoTra -1, CLng(Resto)
            Case LCase("FINALMargenIzquierdoTRA"): DefineFinalMargenIzquierdoTra -1, CLng(Resto)
            Case LCase("FINALMargenSuperiorTRA"): DefineFinalMargenSuperiorTra -1, CLng(Resto)
            Case LCase("FINALMargenInferiorTRA"): DefineFinalMargenInferiorTra -1, CLng(Resto)
        End Select
SIGLINEA:
    Next J
    
    Set FSO = Nothing
End Sub

Public Sub GrabarPackage(sPath As String, Optional AvisarSiExiste As Boolean = True)
    'aqui se graba un paquete con la definicion y un ejemplo de cada una de las imagenes con sus
    'caracteristicas
    'dentro de la carpeta que se graba como parametro tenemos
    'sNamePackage.DEF (que se borra cuando se mete en el paquete)
    'sNamePackage.SKI_ (que contiene el archivo DEF y todas las imagenes)
    
    If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
    
    Dim FSO As New Scripting.FileSystemObject
    
    'el nombre del archivo de definiciones debera tener la extencion DEF para _
        luego en el paquete poder leerlo!
    
    Dim tmpS As String 'todo el contenido del archivo
    'voy cargando mientras voy validando
    
    '************* nombre package *****************************************
    If sNamePackage = "" Then
        MsgBox "No le ha dado un nombre al paquete"
        Exit Sub
    End If
    tmpS = "PACKAGE:" + sNamePackage + Chr(5)
    
    '************* cantidad imagenes *****************************************
    If UBound(sPathImages) = 0 Then
        MsgBox "No hay imagenes elegidas" + vbCrLf + "No se grabara"
        Exit Sub
    End If
    
    '************* nombres imagenes repetidos *****************************************
    Dim C As Long, D As Long
    Dim ThisName As String
    For C = 1 To UBound(sPathImages) - 1
        ThisName = sNameImages(C)
        For D = (C + 1) To UBound(sPathImages)
            If sNameImages(D) = ThisName Then
                MsgBox "Esta repetido el nombre de la imagen" + vbCrLf + _
                    sNameImages(C) + vbCrLf + _
                    "NO se grabara!"
                Exit Sub
            End If
        Next D
    Next C
    
    '************* imagenes *****************************************
    For C = 1 To UBound(sPathImages)
        If FSO.FileExists(sPathImages(C)) = False Then
            MsgBox "No existe la imagen que eligio:" + vbCrLf + sPathImages(C) + vbCrLf + "NO se grabara!"
            Exit Sub
        Else
            If sNameImages(C) = "" Then
                MsgBox "No le puso nombre a la imagen:" + vbCrLf + sPathImages(C) + vbCrLf + "NO se grabara!"
                Exit Sub
            End If
            tmpS = tmpS + "IMAGEPATH:" + sPathImages(C) + Chr(5)
            tmpS = tmpS + "IMAGE:" + sNameImages(C) + Chr(5)
            'despues no sirve mas una vez que esta empaquetado, solo el basename+extencion nada mas
            
            tmpS = tmpS + "MINWIDTH:" + CStr(sMinWidth(C)) + Chr(5)
            tmpS = tmpS + "MAXWIDTH:" + CStr(sMaxWidth(C)) + Chr(5)
            tmpS = tmpS + "MINHEIGHT:" + CStr(sMinHeight(C)) + Chr(5)
            tmpS = tmpS + "MAXHEIGHT:" + CStr(sMaxHeight(C)) + Chr(5)
            
            tmpS = tmpS + "COEF:" + CStr(sCoef(C)) + Chr(5)
            tmpS = tmpS + "TRANSP:" + CStr(CLng(sTransparencia(C))) + Chr(5)
            tmpS = tmpS + "TRANSPDET:" + sTrasnparenciaDescripcion(C) + Chr(5)
            
            tmpS = tmpS + "MinMargenDerechoTrans:" + CStr(sMinMargenDerechoTrans(C)) + Chr(5)
            tmpS = tmpS + "MinMargenIzquierdoTrans:" + CStr(sMinMargenIzquierdoTrans(C)) + Chr(5)
            tmpS = tmpS + "MinMargenSuperiorTrans:" + CStr(sMinMargenSuperiorTrans(C)) + Chr(5)
            tmpS = tmpS + "MinMargenInferiorTrans:" + CStr(sMinMargenInferiorTrans(C)) + Chr(5)
            
            tmpS = tmpS + "MaxMargenDerechoTrans:" + CStr(sMaxMargenDerechoTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenIzquierdoTrans:" + CStr(sMaxMargenIzquierdoTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenSuperiorTrans:" + CStr(sMaxMargenSuperiorTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenInferiorTrans:" + CStr(sMaxMargenInferiorTrans(C)) + Chr(5)
            
        End If
    Next C
    
    If FSO.FileExists(sPath + sNamePackage + ".def") Then
        If AvisarSiExiste Then
            If MsgBox("Ya existe el archivo" + vbCrLf + "¿Desea reemplazarlo?", vbYesNo) = vbYes Then
                FSO.DeleteFile sPath + sNamePackage + ".def"
            Else
                Exit Sub
            End If
        Else
            FSO.DeleteFile sPath + sNamePackage + ".def"
        End If
    End If
    
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(sPath + sNamePackage + ".def", True)
        TE.Write tmpS
    TE.Close
    Set TE = Nothing
    
    'juntar todo en un archivo compilado
    Dim JS As New tbrJUSE.clsJUSE
    JS.Archivo = sPath + sNamePackage + ".SKI_"
    'agrego primero el archivo de definiciones!
    JS.AddFile sPath + sNamePackage + ".def"
    'y cada una de las imagenes!
    For C = 1 To UBound(sPathImages)
        JS.AddFile sPathImages(C)
    Next C
    
    JS.Unir False
    
    FSO.DeleteFile sPath + sNamePackage + ".def", True
    
    Set JS = Nothing
    
    Set FSO = Nothing
End Sub

Public Sub GrabarSKINFromSKI_(sPathSKI_Origen As String, sPathSKIN_Destino As String, _
    Optional AvisarSiExiste As Boolean = True)
    
    Dim FSO As New Scripting.FileSystemObject
    
    'aqui cargo un SKI_ y lo transformo en un skin basico con las imagenes predefinidas
    AbrirSKI_ sPathSKI_Origen
    
    'el destino es sPathSKIN_Destino
    If LCase(Right(sPathSKIN_Destino, 5)) <> ".skin" Then sPathSKIN_Destino = sPathSKIN_Destino + ".SKIN"
    
    
    
    GrabarSKIN sPathSKIN_Destino, False
End Sub

Public Sub GrabarSKIN(sPathSKIN_Destino As String, AvisarSiExiste As Boolean)
    
    Dim FSO As New Scripting.FileSystemObject
    sNameSKIN = FSO.GetBaseName(sPathSKIN_Destino)

    Dim sNameDefN As String 'definion del skin que heredo del .DEF dentro de .SKI_
    sNameDefN = FSO.GetParentFolderName(sPathSKIN_Destino) + "\" + sNameSKIN + ".defN"
    'ahora leyendo sus caraceristicas dejo grabajo el ".SKIN"
    
    Dim tmpS As String 'todo el contenido del archivo
    'voy cargando mientras voy validando
    
    '************* nombre package *****************************************
    tmpS = "SKIN:" + sNameSKIN + Chr(5)
    tmpS = tmpS + "PACKAGE:" + sNamePackage + Chr(5)
    '************* imagenes *****************************************
    Dim C As Long
    For C = 1 To UBound(sPathImages)
        If FSO.FileExists(sPathImages(C)) = False Then
            MsgBox "No existe la imagen !!!" + vbCrLf + sPathImages(C) + vbCrLf + "NO se creara el skin."
            Exit Sub
        Else
            tmpS = tmpS + "IMAGEPATH:" + sPathImages(C) + Chr(5)
            tmpS = tmpS + "IMAGE:" + sNameImages(C) + Chr(5)
            'despues no sirve mas una vez que esta empaquetado, solo el basename+extencion nada mas
            'los datos anteriores los gusrdo para control y como referencia!
            tmpS = tmpS + "MINWIDTH:" + CStr(sMinWidth(C)) + Chr(5)
            tmpS = tmpS + "MAXWIDTH:" + CStr(sMaxWidth(C)) + Chr(5)
            tmpS = tmpS + "FINALWIDTH:" + CStr(sAnchoFinal(C)) + Chr(5)
            
            tmpS = tmpS + "MINHEIGHT:" + CStr(sMinHeight(C)) + Chr(5)
            tmpS = tmpS + "MAXHEIGHT:" + CStr(sMaxHeight(C)) + Chr(5)
            tmpS = tmpS + "FINALHEIGHT:" + CStr(sAltoFinal(C)) + Chr(5)
            
            tmpS = tmpS + "COEF:" + CStr(sCoef(C)) + Chr(5)
            tmpS = tmpS + "TRANSP:" + CStr(CLng(sTransparencia(C))) + Chr(5)
            tmpS = tmpS + "TRANSPDET:" + sTrasnparenciaDescripcion(C) + Chr(5)
            
            tmpS = tmpS + "MinMargenDerechoTrans:" + CStr(sMinMargenDerechoTrans(C)) + Chr(5)
            tmpS = tmpS + "FINALMargenDerechoTRA:" + CStr(sMargenDerechoTransFINAL(C)) + Chr(5)
            
            tmpS = tmpS + "MinMargenIzquierdoTrans:" + CStr(sMinMargenIzquierdoTrans(C)) + Chr(5)
            tmpS = tmpS + "FINALMargenIzquierdoTRA:" + CStr(sMargenIzquierdoTransFINAL(C)) + Chr(5)
            
            tmpS = tmpS + "MinMargenSuperiorTrans:" + CStr(sMinMargenSuperiorTrans(C)) + Chr(5)
            tmpS = tmpS + "FINALMargenSuperiorTRA:" + CStr(sMargenSuperiorTransFINAL(C)) + Chr(5)
            
            tmpS = tmpS + "MinMargenInferiorTrans:" + CStr(sMinMargenInferiorTrans(C)) + Chr(5)
            tmpS = tmpS + "FINALMargenInferiorTRA:" + CStr(sMargenInferiorTransFINAL(C)) + Chr(5)
            
            tmpS = tmpS + "MaxMargenDerechoTrans:" + CStr(sMaxMargenDerechoTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenIzquierdoTrans:" + CStr(sMaxMargenIzquierdoTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenSuperiorTrans:" + CStr(sMaxMargenSuperiorTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenInferiorTrans:" + CStr(sMaxMargenInferiorTrans(C)) + Chr(5)
            
        End If
    Next C
    
    If FSO.FileExists(sNameDefN) Then
        If AvisarSiExiste Then
            If MsgBox("Ya existe el archivo" + vbCrLf + "¿Desea reemplazarlo?", vbYesNo) = vbYes Then
                FSO.DeleteFile sNameDefN
            Else
                Exit Sub
            End If
        Else
            FSO.DeleteFile sNameDefN
        End If
    End If
    
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(sNameDefN, True)
        TE.Write tmpS
    TE.Close
    Set TE = Nothing
    
    'juntar todo en un archivo compilado
    Dim JS As New tbrJUSE.clsJUSE
    JS.Archivo = sPathSKIN_Destino
    'agrego primero el archivo de definiciones!
    JS.AddFile sNameDefN
    'y cada una de las imagenes!
    For C = 1 To UBound(sPathImages)
        JS.AddFile sPathImages(C)
    Next C
    
    JS.Unir False
    FSO.DeleteFile sNameDefN, True
    
    Set JS = Nothing
    Set FSO = Nothing
End Sub

Public Property Get NamePackage() As String
    NamePackage = sNamePackage
End Property

Public Property Let NamePackage(nName As String)
    sNamePackage = nName
End Property

Public Property Get NameSKIN() As String
    NameSKIN = sNameSKIN
End Property

Public Property Let NameSKIN(nName As String)
    sNameSKIN = nName
End Property

Public Sub ChgPathImage(I As Long, valU As String)
    If I > UBound(sPathImages) Then Exit Sub
    If I = -1 Then I = UBound(sPathImages)
    sPathImages(I) = valU
End Sub

Public Sub AddImage(sPath As String)
    CantImgs = CantImgs + 1
    ReDim Preserve sPathImages(CantImgs)
    sPathImages(CantImgs) = sPath 'deberia ver si existe ? NOOOO
    
    ReDim Preserve sNameImages(CantImgs)
    ReDim Preserve sMinWidth(CantImgs): ReDim Preserve sMaxWidth(CantImgs)
    ReDim Preserve sMinHeight(CantImgs): ReDim Preserve sMaxHeight(CantImgs)
    ReDim Preserve sCoef(CantImgs)
    ReDim Preserve sTransparencia(CantImgs)
    ReDim Preserve sTrasnparenciaDescripcion(CantImgs)
    ReDim Preserve sMinMargenDerechoTrans(CantImgs): ReDim Preserve sMinMargenIzquierdoTrans(CantImgs)
    ReDim Preserve sMinMargenSuperiorTrans(CantImgs): ReDim Preserve sMinMargenInferiorTrans(CantImgs)
    ReDim Preserve sMaxMargenDerechoTrans(CantImgs): ReDim Preserve sMaxMargenIzquierdoTrans(CantImgs)
    ReDim Preserve sMaxMargenSuperiorTrans(CantImgs): ReDim Preserve sMaxMargenInferiorTrans(CantImgs)
    ReDim Preserve sAnchoFinal(CantImgs): ReDim Preserve sAltoFinal(CantImgs)
    ReDim Preserve sMargenDerechoTransFINAL(CantImgs): ReDim Preserve sMargenSuperiorTransFINAL(CantImgs)
    ReDim Preserve sMargenIzquierdoTransFINAL(CantImgs): ReDim Preserve sMargenInferiorTransFINAL(CantImgs)
End Sub

Public Sub DefineMinWidth(I As Long, valU As Long)
    If I > UBound(sMinWidth) Then Exit Sub
    If I = -1 Then I = UBound(sMinWidth)
    sMinWidth(I) = valU
End Sub

Public Sub DefineMaxWidth(I As Long, valU As Long)
    If I > UBound(sMaxWidth) Then Exit Sub
    If I = -1 Then I = UBound(sMaxWidth)
    sMaxWidth(I) = valU
End Sub

Public Sub DefineMinHeight(I As Long, valU As Long)
    If I > UBound(sMinHeight) Then Exit Sub
    If I = -1 Then I = UBound(sMinHeight)
    sMinHeight(I) = valU
End Sub

Public Sub DefineMaxHeight(I As Long, valU As Long)
    If I > UBound(sMaxHeight) Then Exit Sub
    If I = -1 Then I = UBound(sMaxHeight)
    sMaxHeight(I) = valU
End Sub

Public Sub DefineCoef(I As Long, valU As Single)
    If I > UBound(sCoef) Then Exit Sub
    If I = -1 Then I = UBound(sCoef)
    sCoef(I) = valU
End Sub

Public Sub DefineTranspType(I As Long, valU As sTRANS)
    If I > UBound(sTransparencia) Then Exit Sub
    If I = -1 Then I = UBound(sTransparencia)
    sTransparencia(I) = valU
End Sub

Public Sub DefineNameImage(I As Long, valU As String)
    If I > UBound(sNameImages) Then Exit Sub
    '-1 es en el ultimo (cuando estoy cargando viene joia
    If I = -1 Then I = UBound(sNameImages)
    sNameImages(I) = valU
End Sub

Public Sub DefineTranspDescripcion(I As Long, valU As String)
    If I > UBound(sTrasnparenciaDescripcion) Then Exit Sub
    If I = -1 Then I = UBound(sTrasnparenciaDescripcion)
    sTrasnparenciaDescripcion(I) = valU
End Sub

Public Sub DefineMinMargenDerechoTrans(I As Long, valU As Long)
    If I > UBound(sMinMargenDerechoTrans) Then Exit Sub
    If I = -1 Then I = UBound(sMinMargenDerechoTrans)
    sMinMargenDerechoTrans(I) = valU
End Sub

Public Sub DefineMinMargenIzquierdoTrans(I As Long, valU As Long)
    If I > UBound(sMinMargenIzquierdoTrans) Then Exit Sub
    If I = -1 Then I = UBound(sMinMargenIzquierdoTrans)
    sMinMargenIzquierdoTrans(I) = valU
End Sub

Public Sub DefineMinMargenSuperiorTrans(I As Long, valU As Long)
    If I > UBound(sMinMargenSuperiorTrans) Then Exit Sub
    If I = -1 Then I = UBound(sMinMargenSuperiorTrans)
    sMinMargenSuperiorTrans(I) = valU
End Sub

Public Sub DefineMinMargenInferiorTrans(I As Long, valU As Long)
    If I > UBound(sMinMargenInferiorTrans) Then Exit Sub
    If I = -1 Then I = UBound(sMinMargenInferiorTrans)
    sMinMargenInferiorTrans(I) = valU
End Sub

Public Sub DefineMaxMargenDerechoTrans(I As Long, valU As Long)
    If I > UBound(sMaxMargenDerechoTrans) Then Exit Sub
    If I = -1 Then I = UBound(sMaxMargenDerechoTrans)
    sMaxMargenDerechoTrans(I) = valU
End Sub

Public Sub DefineMaxMargenIzquierdoTrans(I As Long, valU As Long)
    If I > UBound(sMaxMargenIzquierdoTrans) Then Exit Sub
    If I = -1 Then I = UBound(sMaxMargenIzquierdoTrans)
    sMaxMargenIzquierdoTrans(I) = valU
End Sub

Public Sub DefineMaxMargenSuperiorTrans(I As Long, valU As Long)
    If I > UBound(sMaxMargenSuperiorTrans) Then Exit Sub
    If I = -1 Then I = UBound(sMaxMargenSuperiorTrans)
    sMaxMargenSuperiorTrans(I) = valU
End Sub

Public Sub DefineMaxMargenInferiorTrans(I As Long, valU As Long)
    If I > UBound(sMaxMargenInferiorTrans) Then Exit Sub
    If I = -1 Then I = UBound(sMaxMargenInferiorTrans)
    sMaxMargenInferiorTrans(I) = valU
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function GetpathImage(I As Long) As String
    If I > UBound(sPathImages) Then Exit Function
    GetpathImage = sPathImages(I)
End Function

Public Function GetMinWidth(I As Long) As Long
    If I > UBound(sMinWidth) Then Exit Function
    GetMinWidth = sMinWidth(I)
End Function

Public Function GetMaxWidth(I As Long) As Long
    If I > UBound(sMaxWidth) Then Exit Function
    GetMaxWidth = sMaxWidth(I)
End Function

Public Function GetMinHeight(I As Long) As Long
    If I > UBound(sMinHeight) Then Exit Function
    GetMinHeight = sMinHeight(I)
End Function

Public Function GetMaxHeight(I As Long) As Long
    If I > UBound(sMaxHeight) Then Exit Function
    GetMaxHeight = sMaxHeight(I)
End Function

Public Function GetCoef(I As Long) As Single
    If I > UBound(sCoef) Then Exit Function
    GetCoef = sCoef(I)
End Function

Public Function GetTranspType(I As Long) As sTRANS
    If I > UBound(sTransparencia) Then Exit Function
    If sTransparencia(I) = 0 Then
        GetTranspType = Opcional
    Else
        GetTranspType = sTransparencia(I)
    End If
End Function

Public Function GetNameImage(I As Long) As String
    If I > UBound(sNameImages) Then Exit Function
    GetNameImage = sNameImages(I)
End Function

Public Function GetTranspDescripcion(I As Long) As String
    If I > UBound(sTrasnparenciaDescripcion) Then Exit Function
    GetTranspDescripcion = sTrasnparenciaDescripcion(I)
End Function

Public Function GetMinMargenDerechoTrans(I As Long) As Long
    If I > UBound(sMinMargenDerechoTrans) Then Exit Function
    GetMinMargenDerechoTrans = sMinMargenDerechoTrans(I)
End Function

Public Function GetMinMargenIzquierdoTrans(I As Long) As Long
    If I > UBound(sMinMargenIzquierdoTrans) Then Exit Function
    GetMinMargenIzquierdoTrans = sMinMargenIzquierdoTrans(I)
End Function

Public Function GetMinMargenSuperiorTrans(I As Long) As Long
    If I > UBound(sMinMargenSuperiorTrans) Then Exit Function
    GetMinMargenSuperiorTrans = sMinMargenSuperiorTrans(I)
End Function

Public Function GetMinMargenInferiorTrans(I As Long) As Long
    If I > UBound(sMinMargenInferiorTrans) Then Exit Function
    GetMinMargenInferiorTrans = sMinMargenInferiorTrans(I)
End Function

Public Function GetMaxMargenDerechoTrans(I As Long) As Long
    If I > UBound(sMaxMargenDerechoTrans) Then Exit Function
    GetMaxMargenDerechoTrans = sMaxMargenDerechoTrans(I)
End Function

Public Function GetMaxMargenIzquierdoTrans(I As Long) As Long
    If I > UBound(sMaxMargenIzquierdoTrans) Then Exit Function
    GetMaxMargenIzquierdoTrans = sMaxMargenIzquierdoTrans(I)
End Function

Public Function GetMaxMargenSuperiorTrans(I As Long) As Long
    If I > UBound(sMaxMargenSuperiorTrans) Then Exit Function
    GetMaxMargenSuperiorTrans = sMaxMargenSuperiorTrans(I)
End Function

Public Function GetMaxMargenInferiorTrans(I As Long) As Long
    If I > UBound(sMaxMargenInferiorTrans) Then Exit Function
    GetMaxMargenInferiorTrans = sMaxMargenInferiorTrans(I)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub DefineFinalHeight(I As Long, valU As Long)
    If I > UBound(sAltoFinal) Then Exit Sub
    If I = -1 Then I = UBound(sAltoFinal)
    sAltoFinal(I) = valU
End Sub

Public Sub DefineFinalWidth(I As Long, valU As Long)
    If I > UBound(sAnchoFinal) Then Exit Sub
    If I = -1 Then I = UBound(sAnchoFinal)
    sAnchoFinal(I) = valU
End Sub

Public Sub DefineFinalMargenDerechoTra(I As Long, valU As Long)
    If I > UBound(sMargenDerechoTransFINAL) Then Exit Sub
    If I = -1 Then I = UBound(sMargenDerechoTransFINAL)
    sMargenDerechoTransFINAL(I) = valU
End Sub

Public Sub DefineFinalMargenIzquierdoTra(I As Long, valU As Long)
    If I > UBound(sMargenIzquierdoTransFINAL) Then Exit Sub
    If I = -1 Then I = UBound(sMargenIzquierdoTransFINAL)
    sMargenIzquierdoTransFINAL(I) = valU
End Sub

Public Sub DefineFinalMargenSuperiorTra(I As Long, valU As Long)
    If I > UBound(sMargenSuperiorTransFINAL) Then Exit Sub
    If I = -1 Then I = UBound(sMargenSuperiorTransFINAL)
    sMargenSuperiorTransFINAL(I) = valU
End Sub

Public Sub DefineFinalMargenInferiorTra(I As Long, valU As Long)
    If I > UBound(sMargenInferiorTransFINAL) Then Exit Sub
    If I = -1 Then I = UBound(sMargenInferiorTransFINAL)
    sMargenInferiorTransFINAL(I) = valU
End Sub

'******************+++++++++*************************

Public Function GetFinalHeight(I As Long) As Long
    If I > UBound(sAltoFinal) Then Exit Function
    GetFinalHeight = sAltoFinal(I)
End Function

Public Function GetFinalWidth(I As Long) As Long
    If I > UBound(sAnchoFinal) Then Exit Function
    GetFinalWidth = sAnchoFinal(I)
End Function

Public Function GetFinalMargenDerechoTra(I As Long) As Long
    If I > UBound(sMargenDerechoTransFINAL) Then Exit Function
    GetFinalMargenDerechoTra = sMargenDerechoTransFINAL(I)
End Function

Public Function GetFinalMargenIzquierdoTra(I As Long) As Long
    If I > UBound(sMargenIzquierdoTransFINAL) Then Exit Function
    GetFinalMargenIzquierdoTra = sMargenIzquierdoTransFINAL(I)
End Function

Public Function GetFinalMargenSuperiorTra(I As Long) As Long
    If I > UBound(sMargenSuperiorTransFINAL) Then Exit Function
    GetFinalMargenSuperiorTra = sMargenSuperiorTransFINAL(I)
End Function

Public Function GetFinalMargenInferiorTra(I As Long) As Long
    If I > UBound(sMargenInferiorTransFINAL) Then Exit Function
    GetFinalMargenInferiorTra = sMargenInferiorTransFINAL(I)
End Function

Private Sub Class_Initialize()
    TodoACero
End Sub

Private Sub TodoACero()
    CantImgs = 0
    ReDim Preserve sPathImages(0)
    ReDim Preserve sMinWidth(0)
    ReDim Preserve sMaxWidth(0)
    ReDim Preserve sMinHeight(0)
    ReDim Preserve sMaxHeight(0)
    ReDim Preserve sCoef(0)
    ReDim Preserve sTransparencia(0): sTransparencia(0) = Opcional
    ReDim Preserve sTrasnparenciaDescripcion(0)
    ReDim Preserve sMinMargenDerechoTrans(0)
    ReDim Preserve sMaxMargenDerechoTrans(0)
    ReDim Preserve sMinMargenIzquierdoTrans(0)
    ReDim Preserve sMaxMargenIzquierdoTrans(0)
    ReDim Preserve sMinMargenSuperiorTrans(0)
    ReDim Preserve sMaxMargenSuperiorTrans(0)
    ReDim Preserve sMinMargenInferiorTrans(0)
    ReDim Preserve sMaxMargenInferiorTrans(0)
    ReDim Preserve sAnchoFinal(0): ReDim Preserve sAltoFinal(0)
    ReDim Preserve sMargenDerechoTransFINAL(0): ReDim Preserve sMargenSuperiorTransFINAL(0)
    ReDim Preserve sMargenIzquierdoTransFINAL(0): ReDim Preserve sMargenInferiorTransFINAL(0)
End Sub

Public Sub RemoveImage(I As Long)
    If I > UBound(sPathImages) Then Exit Sub
    
    If CantImgs = 0 Then Exit Sub
    CantImgs = CantImgs - 1
    
    Dim J As Long
    For J = I To UBound(sPathImages) - 1
        sPathImages(I) = sPathImages(I + 1)
        sMinWidth(I) = sMinWidth(I + 1)
        sMaxWidth(I) = sMaxWidth(I + 1)
        sMinHeight(I) = sMinHeight(I + 1)
        sMaxHeight(I) = sMaxHeight(I + 1)
        sCoef(I) = sCoef(I + 1)
        sTransparencia(I) = sTransparencia(I + 1)
        sTrasnparenciaDescripcion(I) = sTrasnparenciaDescripcion(I + 1)
        sMinMargenDerechoTrans(I) = sMinMargenDerechoTrans(I + 1)
        sMaxMargenDerechoTrans(I) = sMaxMargenDerechoTrans(I + 1)
        sMinMargenIzquierdoTrans(I) = sMinMargenIzquierdoTrans(I + 1)
        sMaxMargenIzquierdoTrans(I) = sMaxMargenIzquierdoTrans(I + 1)
        sMinMargenSuperiorTrans(I) = sMinMargenSuperiorTrans(I + 1)
        sMaxMargenSuperiorTrans(I) = sMaxMargenSuperiorTrans(I + 1)
        sMinMargenInferiorTrans(I) = sMinMargenInferiorTrans(I + 1)
        sMaxMargenInferiorTrans(I) = sMaxMargenInferiorTrans(I + 1)
        sAnchoFinal(I) = sAnchoFinal(I + 1)
        sAltoFinal(I) = sAltoFinal(I + 1)
        sMargenDerechoTransFINAL(I) = sMargenDerechoTransFINAL(I + 1)
        sMargenSuperiorTransFINAL(I) = sMargenSuperiorTransFINAL(I + 1)
        sMargenIzquierdoTransFINAL(I) = sMargenIzquierdoTransFINAL(I + 1)
        sMargenInferiorTransFINAL(I) = sMargenInferiorTransFINAL(I + 1)

    Next J
    
    J = UBound(sPathImages) - 1
    
    ReDim Preserve sPathImages(J)
    ReDim Preserve sMinWidth(J)
    ReDim Preserve sMaxWidth(J)
    ReDim Preserve sMinHeight(J)
    ReDim Preserve sMaxHeight(J)
    ReDim Preserve sCoef(J)
    ReDim Preserve sTransparencia(J): sTransparencia(J) = Opcional
    ReDim Preserve sTrasnparenciaDescripcion(J)
    ReDim Preserve sMinMargenDerechoTrans(J)
    ReDim Preserve sMaxMargenDerechoTrans(J)
    ReDim Preserve sMinMargenIzquierdoTrans(J)
    ReDim Preserve sMaxMargenIzquierdoTrans(J)
    ReDim Preserve sMinMargenSuperiorTrans(J)
    ReDim Preserve sMaxMargenSuperiorTrans(J)
    ReDim Preserve sMinMargenInferiorTrans(J)
    ReDim Preserve sMaxMargenInferiorTrans(J)
    ReDim Preserve sAnchoFinal(J): ReDim Preserve sAltoFinal(J)
    ReDim Preserve sMargenDerechoTransFINAL(J): ReDim Preserve sMargenSuperiorTransFINAL(J)
    ReDim Preserve sMargenIzquierdoTransFINAL(J): ReDim Preserve sMargenInferiorTransFINAL(J)
End Sub
