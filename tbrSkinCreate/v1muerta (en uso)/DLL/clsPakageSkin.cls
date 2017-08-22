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
Attribute VB_Exposed = True
Option Explicit

Private sNamePackage As String 'nombre del paquete con las definiciones
Private sNameSKIN As String 'nombre de un skin especifico

Private ArchLOG As String 'se define desde el ejecutable
Private DF As New clsDef

Public Function getDef() As clsDef
    Set getDef = DF 'devuelve todas las propiedades necesarias
End Function

Public Function AbrirSKI_(ArchSKI_ As String) As Long
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
    
    Dim ADN As String 'archivo def n
    ADN = ET + FSO.GetBaseName(ArchSKI_) + ".def"
    If FSO.FileExists(ADN) Then
        loadDef ADN
    Else
        AbrirSKI_ = 1 'alguien le cambio el nombre al skin
    End If
    
    AbrirSKI_ = 0
    Set FSO = Nothing
End Function

Public Function AbrirSKIN(ArchSKIN As String) As Long
    On Local Error GoTo ErrAbrir
    
    TER.Anotar "op0", ArchSKIN
    
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
    TER.Anotar "op1", ET
    If FSO.FolderExists(ET) = False Then FSO.CreateFolder ET
    Dim A As Long
    For A = 1 To JS.CantArchs
        JS.Extract ET, A
        TER.Anotar "op2", A
    Next A
    'cerrar todo
    Set JS = Nothing
    Dim ADN As String 'archivo def n
    ADN = ET + FSO.GetBaseName(ArchSKIN) + ".defN"
    TER.Anotar "op3", ADN
    If FSO.FileExists(ADN) Then
        loadDef ADN
    Else
        AbrirSKIN = 1 'alguien le cambio el nombre al skin
        TER.Anotar "op4"
        Exit Function
    End If
    
    Set FSO = Nothing
    
    AbrirSKIN = 0
    
    Exit Function
    
ErrAbrir:
    TER.AppendLog "ErrOPNE:", TER.ErrToTXT(Err)
    Resume Next
    
End Function

Private Function loadDef(sArch As String) As Long
    On Local Error GoTo errDEFn
    
    TER.Anotar "LD", sArch
    'abre validando que todo exista!
    DF.Clean 'lo dejo en cero
    
    'monto la definicion primero
    Dim FSO As New Scripting.FileSystemObject
    
    'esta todo dentro de la carpeta donde esta el def
    Dim FOL As String
    
    FOL = FSO.GetParentFolderName(sArch)
    If Right(FOL, 1) <> "\" Then FOL = FOL + "\"
    TER.Anotar "LD1", FOL
    Dim TE As TextStream, tmpT As String
    'si se creo con un nombre y se le cambia dara un error por que el DEFN tiene el nombre original que tuvo!
    If FSO.FileExists(sArch) Then
        Set TE = FSO.OpenTextFile(sArch)
            tmpT = TE.ReadAll
        TE.Close
    Else
        TER.Anotar "LD2"
        loadDef = 1
        Exit Function
    End If
    
    Dim Lineas() As String
    Lineas = Split(tmpT, Chr(5))
    Dim J As Long
    Dim Tit() As String, Resto As String 'titulo y resto de cada linea
    For J = 0 To UBound(Lineas)
        TER.Anotar "LD3", Lineas(J)
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
            Case "skin": sNameSKIN = Resto
            'empieza una imagen, agregarla para que se vayan creando los indices!
            Case "imagepath": DF.AddImage FOL + FSO.GetBaseName(Resto) + "." + FSO.GetExtensionName(Resto)
            'id unico de la imagen
            Case "image": DF.DefineNameImage -1, Resto
            Case "minwidth": DF.DefineMinWidth -1, CLng(Resto)
            Case "maxwidth": DF.DefineMaxWidth -1, CLng(Resto)
            
            Case "minheight": DF.DefineMinHeight -1, CLng(Resto)
            Case "maxheight": DF.DefineMaxHeight -1, CLng(Resto)
            Case "coef": DF.DefineCoef -1, CSng(Resto)
            Case "transp": DF.DefineTranspType -1, CLng(Resto)
            Case "transpdet": DF.DefineTranspDescripcion -1, Resto
            
            Case LCase("MinMargenDerechoTrans"): DF.DefineMinMargenDerechoTrans -1, CLng(Resto)
            Case LCase("MinMargenIzquierdoTrans"): DF.DefineMinMargenIzquierdoTrans -1, CLng(Resto)
            Case LCase("MinMargenSuperiorTrans"): DF.DefineMinMargenSuperiorTrans -1, CLng(Resto)
            Case LCase("MinMargenInferiorTrans"): DF.DefineMinMargenInferiorTrans -1, CLng(Resto)
            
            Case LCase("MaxMargenDerechoTrans"): DF.DefineMaxMargenDerechoTrans -1, CLng(Resto)
            Case LCase("MaxMargenIzquierdoTrans"): DF.DefineMaxMargenIzquierdoTrans -1, CLng(Resto)
            Case LCase("MaxMargenSuperiorTrans"): DF.DefineMaxMargenSuperiorTrans -1, CLng(Resto)
            Case LCase("MaxMargenInferiorTrans"): DF.DefineMaxMargenInferiorTrans -1, CLng(Resto)
        
            'solo para skins!
            Case LCase("FINALHEIGHT"): DF.DefineFinalHeight -1, CLng(Resto)
            Case LCase("FINALWIDTH"): DF.DefineFinalWidth -1, CLng(Resto)
            Case LCase("FINALMargenDerechoTRA")
                DF.DefineFinalMargenDerechoTra -1, CLng(Resto)
            Case LCase("FINALMargenIzquierdoTRA")
                DF.DefineFinalMargenIzquierdoTra -1, CLng(Resto)
            Case LCase("FINALMargenSuperiorTRA")
                DF.DefineFinalMargenSuperiorTra -1, CLng(Resto)
            Case LCase("FINALMargenInferiorTRA")
                DF.DefineFinalMargenInferiorTra -1, CLng(Resto)
                
            'COLORES
            Case "idcolor"
                DF.AddColorOnlyName Resto
            Case "lncolor"
                DF.DefineColorValue -1, CLng(Resto)
            
            Case Else
                TER.AppendSinHist "LoadDefNoExiste!:" + vbCrLf + Tit(0)
                
        End Select
SIGLINEA:
    Next J
    
    Set FSO = Nothing
    
    Exit Function
    
errDEFn:
    TER.AppendLog "ErrDEFN:", TER.ErrToTXT(Err)
    Resume Next
    
End Function

Public Sub GrabarPackage(sPath As String, Optional AvisarSiExiste As Boolean = True)
    'aqui se graba un paquete con la definicion y un ejemplo de cada una de las imagenes con sus
    'caracteristicas
    'dentro de la carpeta que se graba como parametro tenemos
    'sNamePackage.DEF (que se borra cuando se mete en el paquete)
    'sNamePackage.SKI_ (que contiene el archivo DEF y todas las imagenes)
    
    Dim FSO As New Scripting.FileSystemObject
    sNamePackage = FSO.GetBaseName(sPath)
    
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
    If DF.GetCantImgs = 0 Then
        MsgBox "No hay imagenes elegidas" + vbCrLf + "No se grabara"
        Exit Sub
    End If
    
    '************* nombres imagenes repetidos *****************************************
    Dim C As Long, D As Long
    Dim ThisName As String
    For C = 1 To DF.GetCantImgs - 1
        ThisName = DF.GetNameImage(C)
        For D = (C + 1) To DF.GetCantImgs
            If LCase(DF.GetNameImage(D)) = LCase(ThisName) Then
                MsgBox "Esta repetido el nombre de la imagen" + vbCrLf + _
                    DF.GetNameImage(C) + vbCrLf + _
                    "NO se grabara!"
                Exit Sub
            End If
        Next D
    Next C
    '************* nombres colores repetidos *****************************************
    
    For C = 1 To DF.GetCantColores - 1
        ThisName = DF.getNameColor(C)
        For D = (C + 1) To DF.GetCantColores
            If LCase(DF.getNameColor(D)) = LCase(ThisName) Then
                MsgBox "Esta repetido el nombre del color" + vbCrLf + _
                    DF.getNameColor(C) + vbCrLf + _
                    "NO se grabara!"
                Exit Sub
            End If
        Next D
    Next C
    '************* imagenes *****************************************
    For C = 1 To DF.GetCantImgs
        If FSO.FileExists(DF.GetpathImage(C)) = False Then
            MsgBox "No existe la imagen que eligio:" + vbCrLf + DF.GetpathImage(C) + vbCrLf + "NO se grabara!"
            Exit Sub
        Else
            If DF.GetNameImage(C) = "" Then
                MsgBox "No le puso nombre a la imagen:" + vbCrLf + DF.GetpathImage(C) + vbCrLf + "NO se grabara!"
                Exit Sub
            End If
            tmpS = tmpS + "IMAGEPATH:" + DF.GetpathImage(C) + Chr(5)
            tmpS = tmpS + "IMAGE:" + DF.GetNameImage(C) + Chr(5)
            'despues no sirve mas una vez que esta empaquetado, solo el basename+extencion nada mas
            
            tmpS = tmpS + "MINWIDTH:" + CStr(DF.GetMinWidth(C)) + Chr(5)
            tmpS = tmpS + "MAXWIDTH:" + CStr(DF.GetMaxWidth(C)) + Chr(5)
            tmpS = tmpS + "MINHEIGHT:" + CStr(DF.GetMinHeight(C)) + Chr(5)
            tmpS = tmpS + "MAXHEIGHT:" + CStr(DF.GetMaxHeight(C)) + Chr(5)
            
            tmpS = tmpS + "COEF:" + CStr(DF.GetCoef(C)) + Chr(5)
            tmpS = tmpS + "TRANSP:" + CStr(CLng(DF.GetTranspType(C))) + Chr(5)
            tmpS = tmpS + "TRANSPDET:" + DF.GetTranspDescripcion(C) + Chr(5)
            
            tmpS = tmpS + "MinMargenDerechoTrans:" + CStr(DF.GetMinMargenDerechoTrans(C)) + Chr(5)
            tmpS = tmpS + "MinMargenIzquierdoTrans:" + CStr(DF.GetMinMargenIzquierdoTrans(C)) + Chr(5)
            tmpS = tmpS + "MinMargenSuperiorTrans:" + CStr(DF.GetMinMargenSuperiorTrans(C)) + Chr(5)
            tmpS = tmpS + "MinMargenInferiorTrans:" + CStr(DF.GetMinMargenInferiorTrans(C)) + Chr(5)
            
            tmpS = tmpS + "MaxMargenDerechoTrans:" + CStr(DF.GetMaxMargenDerechoTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenIzquierdoTrans:" + CStr(DF.GetMaxMargenIzquierdoTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenSuperiorTrans:" + CStr(DF.GetMaxMargenSuperiorTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenInferiorTrans:" + CStr(DF.GetMaxMargenInferiorTrans(C)) + Chr(5)
            
        End If
    Next C
    
    '************* colores *****************************************
    For C = 1 To DF.GetCantColores
        If DF.getNameColor(C) = "" Then
            MsgBox "No le puso nombre al color" + vbCrLf + "NO se grabara!"
            Exit Sub
        End If
        
        'SI O SI PRIMERO ELO NOMBRE! para leerlo ok despues
        tmpS = tmpS + "idColor:" + DF.getNameColor(C) + Chr(5)
        tmpS = tmpS + "lnColor:" + CStr(DF.getColorById(C)) + Chr(5)

    Next C
    
    If FSO.FileExists(FSO.GetParentFolderName(sPath) + "\" + sNamePackage + ".def") Then
        If AvisarSiExiste Then
            If MsgBox("Ya existe el archivo" + vbCrLf + "¿Desea reemplazarlo?", vbYesNo) = vbYes Then
                FSO.DeleteFile FSO.GetParentFolderName(sPath) + "\" + sNamePackage + ".def"
            Else
                Exit Sub
            End If
        Else
            FSO.DeleteFile FSO.GetParentFolderName(sPath) + "\" + sNamePackage + ".def"
        End If
    End If
    
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(FSO.GetParentFolderName(sPath) + "\" + sNamePackage + ".def", True)
        TE.Write tmpS
    TE.Close
    Set TE = Nothing
    
    'juntar todo en un archivo compilado
    Dim JS As New tbrJUSE.clsJUSE
    JS.Archivo = sPath
    'agrego primero el archivo de definiciones!
    JS.AddFile FSO.GetParentFolderName(sPath) + "\" + sNamePackage + ".def"
    'y cada una de las imagenes!
    For C = 1 To DF.GetCantImgs
        JS.AddFile DF.GetpathImage(C)
    Next C
    
    JS.Unir False
    
    FSO.DeleteFile FSO.GetParentFolderName(sPath) + "\" + sNamePackage + ".def"
    
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
    For C = 1 To DF.GetCantImgs
        If FSO.FileExists(DF.GetpathImage(C)) = False Then
            MsgBox "No existe la imagen !!!" + vbCrLf + DF.GetpathImage(C) + vbCrLf + "NO se creara el skin."
            Exit Sub
        Else
            tmpS = tmpS + "IMAGEPATH:" + DF.GetpathImage(C) + Chr(5)
            tmpS = tmpS + "IMAGE:" + DF.GetNameImage(C) + Chr(5)
            'despues no sirve mas una vez que esta empaquetado, solo el basename+extencion nada mas
            
            tmpS = tmpS + "MINWIDTH:" + CStr(DF.GetMinWidth(C)) + Chr(5)
            tmpS = tmpS + "MAXWIDTH:" + CStr(DF.GetMaxWidth(C)) + Chr(5)
            tmpS = tmpS + "FINALWIDTH:" + CStr(DF.GetFinalWidth(C)) + Chr(5)
            
            tmpS = tmpS + "MINHEIGHT:" + CStr(DF.GetMinHeight(C)) + Chr(5)
            tmpS = tmpS + "MAXHEIGHT:" + CStr(DF.GetMaxHeight(C)) + Chr(5)
            tmpS = tmpS + "FINALHEIGHT:" + CStr(DF.GetFinalHeight(C)) + Chr(5)
            
            tmpS = tmpS + "COEF:" + CStr(DF.GetCoef(C)) + Chr(5)
            tmpS = tmpS + "TRANSP:" + CStr(CLng(DF.GetTranspType(C))) + Chr(5)
            tmpS = tmpS + "TRANSPDET:" + DF.GetTranspDescripcion(C) + Chr(5)
            
            tmpS = tmpS + "MinMargenDerechoTrans:" + CStr(DF.GetMinMargenDerechoTrans(C)) + Chr(5)
            tmpS = tmpS + "FINALMargenDerechoTRA:" + CStr(DF.GetFinalMargenDerechoTra(C)) + Chr(5)
            
            tmpS = tmpS + "MinMargenIzquierdoTrans:" + CStr(DF.GetMinMargenIzquierdoTrans(C)) + Chr(5)
            tmpS = tmpS + "FINALMargenIzquierdoTRA:" + CStr(DF.GetFinalMargenIzquierdoTra(C)) + Chr(5)
            
            tmpS = tmpS + "MinMargenSuperiorTrans:" + CStr(DF.GetMinMargenSuperiorTrans(C)) + Chr(5)
            tmpS = tmpS + "FINALMargenSuperiorTRA:" + CStr(DF.GetFinalMargenSuperiorTra(C)) + Chr(5)
            
            tmpS = tmpS + "MinMargenInferiorTrans:" + CStr(DF.GetMinMargenInferiorTrans(C)) + Chr(5)
            tmpS = tmpS + "FINALMargenInferiorTRA:" + CStr(DF.GetFinalMargenInferiorTra(C)) + Chr(5)
            
            tmpS = tmpS + "MaxMargenDerechoTrans:" + CStr(DF.GetMaxMargenDerechoTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenIzquierdoTrans:" + CStr(DF.GetMaxMargenIzquierdoTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenSuperiorTrans:" + CStr(DF.GetMaxMargenSuperiorTrans(C)) + Chr(5)
            tmpS = tmpS + "MaxMargenInferiorTrans:" + CStr(DF.GetMaxMargenInferiorTrans(C)) + Chr(5)
            
        End If
    Next C
        
    'agregar colores que ya estaban en SKI!!! 30/10/2010
    
    For C = 1 To DF.GetCantColores
        If DF.getNameColor(C) = "" Then
            MsgBox "No le puso nombre al color" + vbCrLf + "NO se grabara!"
            Exit Sub
        End If
        
        'SI O SI PRIMERO ELO NOMBRE! para leerlo ok despues
        tmpS = tmpS + "idColor:" + DF.getNameColor(C) + Chr(5)
        tmpS = tmpS + "lnColor:" + CStr(DF.getColorById(C)) + Chr(5)

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
    For C = 1 To DF.GetCantImgs
        JS.AddFile DF.GetpathImage(C)
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

Private Sub Class_Initialize()
    'valor predeterminado
    ArchLOG = App.Path + "\regSKIN.log"
    TER.FileLog = ArchLOG
    TER.LargoAcumula = 900
    
    DF.Clean
End Sub

Public Sub SetLogErr(Arch As String)
    ArchLOG = Arch
    TER.FileLog = ArchLOG
End Sub
