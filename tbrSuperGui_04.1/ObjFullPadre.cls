VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjFullPadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cada uno de los pictureboxes contenedores de los objetos gráficos
Public sName As String
Public sHDC As Long
Public pathImgFondo As String 'imagen de fondo si usa

'el ancho y el alto luego necesitaran ser transformados de la misma forma que los SGO
'SEGUIRAQUI posiblemente este deberia ser otro objetos SGO
'pero los SGO no tiene imagen de fondo
'estos objPadres podrian ser un objFull agregandoles la propiedad hdc ...
'no tengo ganas ...
'ya se ... le pongo un objetoSGO oculto para que me mida las coordenadas
Private mCoord As New tbrSuperGuiObject
'je ... soy un genio

Private fIMG As New clsFullImgs 'para la imagen de fondo y el archivo de cfg y otras cosas que puedan surgir en el futuro

'lista de sus hijos
Private OBS() As New objFULL  ' As New tbrSuperGuiObject 'OBjetoS

'me interesa poder refrescarme yo mismo sin depender de escribir en el codigo del ejecutable
Private mPictureBoxUsado As Object
'todavia no usado (nov09)

Private mPath As String 'archivo que se grabo/leyo

'para que solo se abra una sola vez el GuiEdit uso un obeto por cada objFull
Dim F_edit As frmEditPadre

Public Function GetSgoInterno() As tbrSuperGuiObject
    Set GetSgoInterno = mCoord
End Function

Public Function Save(Optional sFile As String = "") As Long
    On Local Error GoTo ErrSaveFull
    terr.Anotar "skin080", sFile, sName
    
    
    If sFile = "" Then 'quiere regrabar el mismo
        If fso.FileExists(mPath) = False Then
            terr.AppendLog "TrySaveOBJ__PADRE__Fail!!"
            Exit Function
        Else
            sFile = mPath
        End If
    End If

    'hacer un archivo de texto con toda la config
    'uso el tbrConfig

    Dim C5 As New tbrConfig05.clstbrCfg05
    Dim fCfg As String
    fCfg = tmpFolder + "cfg_PADRE.txt"
    'limpiarlo si existe
    If fso.FileExists(fCfg) Then fso.DeleteFile fCfg, True
    
    terr.Anotar "skin081", fCfg
    
    C5.SetMainFile fCfg
    
    '******************************************
    'estos son calculados y asignados por el SGO
    C5.SetV "sName", sName
    'coordenadas internas
    C5.SetV "X", mCoord.X
    C5.SetV "Y", mCoord.Y
    C5.SetV "W", mCoord.W
    C5.SetV "H", mCoord.H
    
    '******************************************
    
    C5.GrabarConfig
    
    terr.Anotar "skin082"
    
    'empaquetar todo
    Dim F As clsFullImgs_FILE
    Set F = fIMG.AddFileByPath(fCfg, , "BASE_CONFIG_PADRE") 'config de texto
    F.sID = "BASE_CONFIG_PADRE"
    
    terr.Anotar "skin083"
    'ver si usa imagen de fondo
    If fso.FileExists(pathImgFondo) Then
        Dim F2 As clsFullImgs_FILE
        Set F2 = fIMG.AddFileByPath(pathImgFondo, , "FondoPadre")
        F2.sID = "FondoPadre"
        
    End If
    
    terr.Anotar "skin084"
    
    'grabar todos sus objetos hijos
    Dim tmpSGO As objFULL
    Dim M As Long
    Dim fObj As String 'path de cada hijo al grabarse
    
    For M = 1 To GetSgoMaxID
        'cada uno se inicializa diferente
        
        Set tmpSGO = GetSgoByID(M)
        fObj = tmpFolder + CStr(M) + ".hijo"
        
        terr.Anotar "skin061-21", tmpSGO.Tipo, fObj
        tmpSGO.Save fObj
        Dim FF As clsFullImgs_FILE
        Set FF = fIMG.AddFileByPath(fObj, , "Obj:" + CStr(M))
        FF.sID = "Obj:" + CStr(M)
        
    Next M
    
    terr.Anotar "skin084b"
    
    fIMG.SaveFile sFile
    
    terr.Anotar "skin086"
    
    mPath = sFile
    Save = 0
    
    Exit Function
    
ErrSaveFull:
    Save = Err.Number
    terr.AppendLog "skn3085", terr.ErrToTXT(Err)
End Function

Public Sub Load(sFile As String, sFolder As String)
    On Local Error GoTo ErrLoadFull
    terr.Anotar "skin087", sFile, sFolder
    
    'descomprimir el archivo
    Dim H As Long
    H = fIMG.OpenSkinFile(sFile, sFolder)
    
    terr.Anotar "skin088", H
    If H > 0 Then 'error al abrirse archivo
        terr.AppendLog "No089OpenSk:" + CStr(H) + ":" + sFile + ":" + sFolder
        Exit Sub
    End If
    
    'buscar entre todos los archivos cual es de la config (empieza con "CFG_")
    Dim FI As clsFullImgs_FILE
    Set FI = fIMG.GetFileByName("cfg_padre")
    
    terr.Anotar "skin090"
    
    'abrir la config y aplicar los valores
    Dim fCfg As String 'archivo de la config
    fCfg = FI.path
    
    terr.Anotar "skin091"
    Dim C5 As New tbrConfig05.clstbrCfg05
    C5.AbrirCfg fCfg
    
    terr.Anotar "skin092", fCfg
    sName = C5.GetV("sName")
    'coordenadas internas
    mCoord.X = C5.GetLong("X")
    mCoord.Y = C5.GetLong("Y")
    mCoord.W = C5.GetLong("W")
    mCoord.H = C5.GetLong("H")
    
    terr.Anotar "skin093", sName
    
    'si no existe queda en "", no es problema (registrara un AppendLog pero despues lo veo) seguiraqui
    pathImgFondo = fIMG.GetFileBysID("FondoPadre", False).path
    
    'buscar y cargar todos los hijos
    Dim M As Long
    For M = 1 To fIMG.GetFileCantidad
        terr.Anotar "skin093-d", M
        Dim Fi2 As clsFullImgs_FILE
        Set Fi2 = fIMG.GetFileById(M)
        If LCase(fso.GetExtensionName(Fi2.path)) = "hijo" Then
            
            Dim FolExt As String
            FolExt = fso.BuildPath(sFolder, CStr(M))
            terr.Anotar "skin093-e", Fi2.path, FolExt
            
            If fso.FolderExists(FolExt) Then fso.DeleteFolder FolExt
            fso.CreateFolder FolExt
            
            
            Dim tmpSGO As New objFULL
            tmpSGO.Load Fi2.path, FolExt
            AppendSGO tmpSGO
        End If
                
    Next M
    
    mPath = sFile
    Exit Sub
    
ErrLoadFull:
    terr.AppendLog "skin040", terr.ErrToTXT(Err)
End Sub

Public Sub AppendSGO(newSGO As objFULL)
    Dim k As Long
    k = UBound(OBS) + 1
    ReDim Preserve OBS(k)
    Set OBS(k) = newSGO
End Sub

Public Function AddSgo(sName As String) As objFULL
    Dim k As Long
    k = UBound(OBS) + 1
    ReDim Preserve OBS(k)
    'Set OBS(k) = New objFULL
    'SEGUIRAQUI deberia verificar que no existan dos con el mismo nombre
    Dim SGO As New tbrSuperGuiObject
    Set OBS(k).oSimple = SGO
    OBS(k).oSimple.SGOName = sName
    
    'el objeto del manu se asigna despues
    
    Set AddSgo = OBS(k)
    
End Function

Public Function GetSgoMaxID() As Long
    GetSgoMaxID = UBound(OBS)
End Function

Public Function GetSgoByID(Ix As Long) As objFULL
    If Ix < 0 Or Ix > UBound(OBS) Then
        Set GetSgoByID = OBS(0)
    Else
        Set GetSgoByID = OBS(Ix)
    End If
End Function

Public Function getSGO(sName As String) As objFULL
    Dim k As Long
    For k = 1 To UBound(OBS)
        If Trim(LCase(OBS(k).oSimple.SGOName)) = Trim(LCase(sName)) Then
            Set getSGO = OBS(k)
            Exit Function
        End If
    Next k
    
    Set getSGO = OBS(0) 'para que no de error doy uno vacio
End Function

Public Sub Clear()
    ReDim OBS(0)
End Sub

'mostrar todo lo que tiene
Public Function toString() As String
    Dim M As Long, TMP As String
    TMP = ""
    For M = 1 To UBound(OBS)
        TMP = TMP + "OBS[" + CStr(M) + "]" + vbCrLf + _
                    OBS(M).oSimple.toString + vbCrLf + _
                    " * Tipo " + CStr(OBS(M).Tipo) + " (" + getStrTipo(OBS(M).Tipo) + ")"
                    
        TMP = TMP + vbCrLf + vbCrLf
    
    Next M
    
    toString = TMP
End Function

Private Sub Class_Initialize()
    Clear
End Sub

'funcon para hacerle algo a TODOS los objetos graficos, inicializar y cerrra por ejemplo
Public Sub INIT_GRAPH(Optional sClose As String = "INIT")

    terr.Anotar "fafk-2", sClose
    On Local Error GoTo ERRGR
    
    'actualizar las corrdenadas de este frm
    mCoord.Traslate
    'y todos sus descendientes
    TraslateALL
    
    '**************************
    'si este padre tiene una imagen de fondo cargarla [solo si estoy mostrando]
    If sClose = "INIT" Then
        Dim H As Long
        H = PintarFondo
        terr.Anotar "fafk-90", H
    End If
    '**************************

    'hacer todos los incializargraficos que correspondan segun indicaciones del manu!
    'se hace aqui por que ya se han traducido las coordenadas en el mismo load
    
    'seguiraqui - necesito los HDC de cada formulario
    'seguiraqui -
    
    Dim M As Long
    
    'todos usan el sgo de cada ID
    Dim tmpSGO As objFULL
    For M = 1 To GetSgoMaxID
        'cada uno se inicializa diferente
        
        Set tmpSGO = GetSgoByID(M)
        
        terr.Anotar "initfg_" + CStr(tmpSGO.Tipo) + " : " + CStr(tmpSGO.oSimple.X2) + " : " + CStr(tmpSGO.oSimple.Y2) + " : " + CStr(tmpSGO.oSimple.W2) + " : " + CStr(tmpSGO.oSimple.H2)
        
        Select Case tmpSGO.Tipo
            
            Case en_clsLabel
                Dim OV_01 As clsLabel
                Set OV_01 = tmpSGO.oManu
                If sClose = "INIT" Then
                    
                    OV_01.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColor"), _
                        tmpSGO.oSimple.GetProp_LNG("FntColorContorno")
                    
                    OV_01.IniciarGraficos sHDC, _
                        tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, tmpSGO.oSimple.H2, _
                        tmpSGO.oSimple.GetProp_BOOL("alphaHabilitado"), tmpSGO.oSimple.GetProp_LNG("AlpahColorLong"), _
                        tmpSGO.oSimple.GetProp_LNG("AlphaCantidad")
                    'por si el tipo grabo algo
                    'OV_01.SetTexto tmpSGO.oSimple.GetProp("TextoActual").ValueProp
                    OV_01.SetTexto tmpSGO.oSimple.GetProp_STR("TextoActual") 'la de arriba tambien deberia ser valida
                    
                    OV_01.Renderizar
                End If
                
                If sClose = "CLOSE" Then
                    OV_01.CerrarGraficos
                End If
            Case en_clsTemasManager '*********LISTAS DE CANCIONES O DE RITMOS O DE LETRAS; ETC
            
                Dim OV_02 As clsTemasManager
                Set OV_02 = tmpSGO.oManu
                
                If sClose = "INIT" Then
                        
                    OV_02.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColorUnSel"), _
                        tmpSGO.oSimple.GetProp_LNG("FontColorSel"), tmpSGO.oSimple.GetProp_LNG("FntColorContorno")
                    
                    OV_02.IniciarGraficos sHDC, _
                        tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, tmpSGO.oSimple.H2, _
                        tmpSGO.oSimple.GetProp_BOOL("AlphaB"), tmpSGO.oSimple.GetProp_LNG("ColorSel")
                    
                    Dim ELM As clsTemaList
                    Dim JJ As Long
                    For JJ = 1 To 7
                        Set ELM = OV_02.AgregarElemento
                        ELM.Titulo = "Titulo " + CStr(JJ)
                        ELM.Duracion = "00:0" + CStr(JJ)
                        ELM.Numero = JJ
                        ELM.TagMisterioso = "tag" + CStr(JJ)
                    Next JJ
                    
                    OV_02.DibujarTexto 0 'creo que el -1 no marca ningun elemento y el cero el primero
                    
                    OV_02.IniciarTouchScreen
                    
                    'no tiene renderizar aparentemente
                    'seguiraqui ver si necesita otra cosa
                    
                End If
                
                If sClose = "CLOSE" Then
                    OV_02.CerrarGraficos
                End If
                
            Case en_tbrTextoSelect '*************listas horizontales ritmos, letras
            
                Dim OV_03 As tbrTextoSelect
                Set OV_03 = tmpSGO.oManu
                
                If sClose = "INIT" Then
                
                    OV_03.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColorUnSel"), _
                        tmpSGO.oSimple.GetProp_LNG("FontColorSel"), tmpSGO.oSimple.GetProp_LNG("FntColorContorno")
                    
                    OV_03.IniciarGraficos sHDC, _
                        tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, tmpSGO.oSimple.H2
                
                    
                    OV_03.AgregarElemento "elem1"
                    OV_03.AgregarElemento "elem2"
                    OV_03.AgregarElemento "elem3"
                    
                    OV_03.DibujarTexto 1, False
                    
                End If
                
                If sClose = "CLOSE" Then
                    OV_03.CerrarGraficos
                End If
                
            Case en_tbrPromociones2
                Dim OV_04 As tbrPromociones2
                Set OV_04 = tmpSGO.oManu
                                
                If sClose = "INIT" Then
                
                    OV_04.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColor")
                        
                    OV_04.IniciarGraficos sHDC, _
                        tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, tmpSGO.oSimple.H2, _
                            tmpSGO.oSimple.GetProp_LNG("PixelSalteo"), tmpSGO.oSimple.GetProp_LNG("TiemposEntreTextos")
                    
                    'las propiedades "texto_num" son despues agregadas como textos
                    'buscar las promos
                    Dim J As Long, Txt As String
                    For J = 1 To 10 'maximo 10 textos
                        Txt = tmpSGO.oSimple.GetProp_STR("Texto_" + CStr(J))
                        If Txt <> "" And Txt <> "0" Then
                            OV_04.AgregarPromo Txt
                        End If
                    Next J
                    
                    
                    
                    
                    'renderizar (necesito que lo hagan todo el tiempo!!!)
                    OV_04.DibujarTexto
                    
                End If
                
                If sClose = "CLOSE" Then
                    OV_04.CerrarGraficos
                End If
            
            Case en_clsDiscoManager
                Dim OV_05 As clsDiscoManager
                Set OV_05 = tmpSGO.oManu
                OV_05.IniciarGraficos sHDC, tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, _
                tmpSGO.oSimple.W2, tmpSGO.oSimple.H2, False, _
                OV_05.MostrarModoNumpad, tmpSGO.oSimple.GetProp_LNG("ColorDiscoSel")
    
                terr.Anotar "aabt30"
                
                If sClose = "INIT" Then
                    OV_05.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColor"), _
                        tmpSGO.oSimple.GetProp_LNG("FntColorContorno")
                
                    OV_05.IniciarPNGs tmpSGO.getFIMG.GetFileBysID("PNGMarcoDisco", True).path
                    
                    OV_05.IniciarGraficos sHDC, tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, _
                        tmpSGO.oSimple.H2, tmpSGO.oSimple.GetProp_BOOL("alphaHabilitado"), False, tmpSGO.oSimple.GetProp_LNG("ColorDiscoSel")
                    
                    
                    OV_05.CleanExtenciones
                    OV_05.AgregarExtencion "mp3"
                    
                    
                    Dim FO As Folder
                    Set FO = fso.GetFolder("C:\mm")
                    Dim Fo2 As Folder
                    For Each Fo2 In FO.SubFolders
                        OV_05.AgregarDisco Fo2.path
                    Next
                    
                    OV_05.MostrarTodo True
                    OV_05.Renderizar
                End If
            
                If sClose = "CLOSE" Then
                    OV_05.CerrarGraficos
                End If
            
            
            'Case en_clsMANU 'NO EXISTE, fue solo de pruebas
            
            Case en_clsNumeriks
            
                Dim OV_06 As clsNumeriks
                Set OV_06 = tmpSGO.oManu
                If sClose = "INIT" Then
                    OV_06.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColor"), _
                        tmpSGO.oSimple.GetProp_LNG("FntColorContorno")
                    
                    OV_06.IniciarGraficos sHDC, _
                        tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, tmpSGO.oSimple.H2, _
                        tmpSGO.oSimple.GetProp_BOOL("alphaHabilitado"), tmpSGO.oSimple.GetProp_LNG("AlpahColorLong"), _
                        tmpSGO.oSimple.GetProp_LNG("AlphaCantidad")
                    
                    OV_06.SetNumero "888" 'especifico de este objeto
                    
                    OV_06.Renderizar
                End If
                
                If sClose = "CLOSE" Then
                    OV_06.CerrarGraficos
                End If
    
            
            Case en_clsPNGBoton 'NO DEFINIDO, una vez usado definir su modo de inicio y cierre
                
                Dim OV_07 As clsPNGBoton
                Set OV_07 = tmpSGO.oManu
                
                If sClose = "INIT" Then
                    
                    OV_07.IniciarPNGs tmpSGO.getFIMG.GetFileBysID("PNGUnSel", True).path
                    
                    'IniciarGraficos(DondeObjetoHdc As Long, XFondo As Long, YFondo As Long)
                    OV_07.IniciarGraficos sHDC, tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2
                    
                    OV_07.DibujarBoton
                    OV_07.Renderizar
                    
                End If
                
                If sClose = "CLOSE" Then
                    OV_07.CerrarGraficos
                End If
                
                
                
                
            Case Else
                terr.AppendSinHist "aaaa:" + tmpSGO.Tipo + " ///TipoDesc:" + sClose
                    
        End Select
    
    Next M
    
    Exit Sub
    
ERRGR:
    terr.AppendLog "ErrGF_G", terr.ErrToTXT(Err)
    Resume Next
    
End Sub

'usoo la imagen de fondo para pintar e informo si no esta o da error!
Public Function PintarFondo() As Long
    On Local Error GoTo noPinta
    'pintar el fondo del formulario si lo tuviera
    If pathImgFondo = "" Then
        ' o no configuro nada o es una falla ...
        PintarFondo = -2
        terr.AppendSinHist "fafk-91"
    Else
        If fso.FileExists(pathImgFondo) Then 'seguiraqui ¿usar imagen predeterminada?
            'seguiraqui se estira de forma predeterminada
            CargarImagenEnHDC pathImgFondo, sHDC
            PintarFondo = 0 'aparentemente ok!
        Else
            PintarFondo = -1
            terr.AppendSinHist "fafk-93"
        End If
    End If
    
    Exit Function
noPinta:
    PintarFondo = Err.Number
    terr.AppendLog "fafk-92", H
End Function

'no estoy tan seguro de que esto vaya aqui o de que deba existir ...
Public Function getStrTipo(sTipo As tipoManu)
    'ACTUALIZAR SI HAY MAS OBJETOS
    Select Case sTipo
        Case en_clsMANU: getStrTipo = "en_clsMANU"
        Case en_clsLabel: getStrTipo = "en_clsLabel"
        Case en_clsNumeriks: getStrTipo = "en_clsNumeriks"
        Case en_clsPNGBoton: getStrTipo = "en_clsPNGBoton"
        Case en_clsTemasManager: getStrTipo = "en_clsTemasManager"
        Case en_tbrTextoSelect: getStrTipo = "en_tbrTextoSelect"
        Case en_tbrPromociones2: getStrTipo = "en_tbrPromociones2"
        Case en_clsDiscoManager: getStrTipo = "en_clsDiscoManager"
    End Select
End Function

'para que el generador de skins pueda mostrar todos los posibles
Public Function CantTiposDatos() As Long
    CantTiposDatos = 8 'ACTUALIZAR SI HAY MAS OBJETOS
End Function

'----------------------------------------------------------------
'REFERENCIA AL PICTUREBOX EXTERNO--------------------------------
' NO IMPLEMENTADO TODAVIA, CUANDO QUIERA HACER REFRESH DESDE ESTA DLL SOLO DEBO REFERENCIAR EL PICTUBOX USADO
'----------------------------------------------------------------
Public Property Let PictureBoxUsado(newObj As Object)
    mPictureBoxUsado = newObj
    sHDC = mPictureBoxUsado.hdc
End Property

Public Property Get PictureBoxUsado() As Object
    PictureBoxUsado = mPictureBoxUsado
End Property

Public Sub Refresh()
    mPictureBoxUsado.Refresh
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
Public Sub TraslateALL()
    Dim k As Long
    For k = 1 To UBound(OBS)
        'QUEDEAQUI definir bien los tamaños teniendo mas en cuenta los formularios que el tamaño de la pantalla
        OBS(k).oSimple.Traslate
    Next k
End Sub

Public Sub GuiEdit()
    If F_edit Is Nothing Then
        Set F_edit = New frmEditPadre
    End If
    F_edit.SetObjPadre Me
    F_edit.Caption = "Padre: " + sName
    F_edit.Show
End Sub
