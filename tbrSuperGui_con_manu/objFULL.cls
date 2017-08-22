VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objFULL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'antes habia una propiedad String (tipo)
Public Enum tipoManu
    en_clsMANU = 1          '"manu"
    en_clsLabel = 2         '"label"
    en_clsNumeriks = 3      '"numerik"
    en_clsPNGBoton = 4      '"pngboton"
    en_clsTemasManager = 5  '"lista"            'coleccion de temasList
    en_tbrTextoSelect = 6   '"listah"
    en_tbrPromociones2 = 7  '"promociones2"     'textos que van cambiando
    en_clsDiscoManager = 8  '"discomanager"
End Enum


'suma de un objeto basico mio y el del manu correspondiente
Public oSimple As tbrSuperGuiObject     'OBjetoS
Public oManu As Object     'varia segun el tipo del manu que sea
Public Tipo As tipoManu
'tipo del objeto del manu
Public Padre As ObjFullPadre
Private FImg As New clsFullImgs

'se necesita asignar ,un objeto del manu a cada objeto de los que administra esta clase
'necesito un proveedor de objetos del manu para cada tipo

Public Function GetNewManu(sTipo As tipoManu) As Object

    terr.Anotar "skin029", sTipo

    Dim x As Object 'valor de retorno
    
    If sTipo = en_clsMANU Then Set x = New clsMANU
    If sTipo = en_clsLabel Then Set x = New clsLabel
    If sTipo = en_clsNumeriks Then Set x = New clsNumeriks
    If sTipo = en_clsPNGBoton Then Set x = New clsPNGBoton
    If sTipo = en_clsTemasManager Then Set x = New clsTemasManager
    If sTipo = en_tbrTextoSelect Then Set x = New tbrTextoSelect
    If sTipo = en_tbrPromociones2 Then Set x = New tbrPromociones2
    If sTipo = en_clsDiscoManager Then Set x = New clsDiscoManager
    
    Set GetNewManu = x
    
    terr.Anotar "skin030"
    
End Function


'hacer un objeto FullImg que contenga todos los archivos que necesite esta clase
'de modo que despues se pueda volver con un load
Public Function Save(sFile As String)

    On Local Error GoTo ErrSaveFull
    terr.Anotar "skin031", sFile

    'parte comun a todos los objetos

    'hacer un archivo de texto con toda la config
    'uso el tbrConfig

    Dim C5 As New tbrConfig05.clstbrCfg05
    Dim fCfg As String
    fCfg = tmpFolder + "cfg_" + CStr(Tipo) + "_" + SGOName + ".txt"
    'limpiarlo si existe
    If fso.FileExists(fCfg) Then fso.DeleteFile fCfg, True
    
    terr.Anotar "skin032", fCfg
    
    C5.SetMainFile fCfg

    'seguiraqui asegurarse sincronizacion entre estas propiedades (oSimple) y las oManu
    UpdatePropiedades 'cargo valores reales del objeto del manu en una estructura mas linda en un tbrSGObject
    
    terr.Anotar "skin033"
    
    '******************************************
    'estos son calculados y asignados por el SGO
    C5.SetV "X", CStr(oSimple.x)
    C5.SetV "Y", CStr(oSimple.y)
    C5.SetV "Ancho", CStr(oSimple.W)
    C5.SetV "Alto", CStr(oSimple.H)
    C5.SetV "Estirable", CStr(CLng(oSimple.Estirable))
    C5.SetV "AlignH", CStr(oSimple.AlignementH)
    C5.SetV "AlignV", CStr(oSimple.AlignementV)
    '******************************************
    
    'las propiedades diferentes de las clases del manu se agregan como "Propis" en oSimple
    'cada una tiene diferentes pero se cargan de un solo modo como colecciones de propis aqui para simplificar todo
    
    Dim H As Long, P As PROPIs
    'SACADO PRO QUE EN SG_2 no se usa y tengo problemas de SVN
    'For H = 1 To oSimple.GetPropCantidad
    '    P = oSimple.GetPropByID(H)
    '    C5.SetV P.NameProp, P.ValueProp
    '    terr.Anotar "skin034", H, P.NameProp, P.ValueProp
    'Next H
    
    terr.Anotar "skin035"
    C5.GrabarConfig
    
    terr.Anotar "skin036"
    
    'empaquetar todo
    Dim F As clsFullImgs_FILE
    Set F = FImg.AddFileByPath(fCfg) 'config de texto
    F.sID = "BASE_CONFIG"
    
    terr.Anotar "skin037"
    'ver si la clase del manu tiene imagenes que guardar
    'pudiera ser posible que todos los objetos del manu tengan un FullImg para respetar una estructura comun
    'de todas formas todavia no parece necesario y lo manjeo segun caso
    
    FImg.AddFilesFromOtherFIMG oManu.GetFIMGs
    
    terr.Anotar "skin038"
    
    FImg.SaveFile sFile
    
    terr.Anotar "skin039"
    Exit Function
    
ErrSaveFull:
    terr.AppendLog "skn3021", terr.ErrToTXT(Err)
    
End Function

'indica que archivo descomprimir y donde grabar los archivos extraidos
Public Function Load(sFile As String, sFolder As String)

    On Local Error GoTo ErrLoadFull
    terr.Anotar "skin044", sFile, sFolder
    
    'descomprimir el archivo
    
    Dim H As Long
    H = FImg.OpenSkinFile(sFile, sFolder)
    
    terr.Anotar "skin045", H
    If H > 0 Then 'error al abrirse archivo
        terr.AppendLog "NoOpenSk:" + CStr(H) + ":" + sFile + ":" + sFolder
        Exit Function
    End If
    
    'buscar entre todos los archivos cual es de la config (empieza con "CFG_")
    Dim FI As clsFullImgs_FILE
    Set FI = FImg.GetFileByName("cfg_*") 'buscar un archivo que empiece con "cfg_" (SOLO DEBE HABER UNO)
    
    terr.Anotar "skin046"
    'ver de que tipo y que nombre tenia
    Dim n As String, SP() As String
    n = FI.BaseName
    SP = Split(n, "_")
    
    terr.Anotar "skin047", n, SP(1), SP(2)
    '***********************************
    Tipo = CLng(SP(1))          'ASIGNAR estos valores importantes
    SGOName = SP(2)
    '***********************************
    
    'abrir la config y aplicar los valores
    Dim fCfg As String 'archivo de la config
    fCfg = FI.path
    
    terr.Anotar "skin048"
    Dim C5 As New tbrConfig05.clstbrCfg05
    C5.AbrirCfg fCfg
    
    terr.Anotar "skin049"
    Dim J As Long, CF As String
    For J = 1 To C5.GetConfigCantidad
        
        CF = LCase(C5.GetNameCfgByIndex(J))
        terr.Anotar "skin050", J, CF
        Select Case CF
        '******************************************
        'estos son calculados y asignados por el SGO
            Case "x"
                oSimple.x = C5.GetLong("X")
            Case "y"
                oSimple.y = C5.GetLong("Y")
            Case "ancho"
                oSimple.W = C5.GetLong("Ancho")
            Case "alto"
                oSimple.H = C5.GetLong("Alto")
            Case "estirable"
                oSimple.Estirable = C5.GetBool("Estirable")
            Case "alignh"
                oSimple.AlignementH = C5.GetLong("AlignH")
            Case "alignv"
                oSimple.AlignementV = C5.GetLong("AlignV")
        
        'los otros
            
            Case Else
                oSimple.AddProp CF, C5.GetV(CF)
        
        End Select
        
        '******************************************
        
    Next J
    
    terr.Anotar "skin051"
    'LAS IMAGENES O ARCHIVOS EXTERNOS QUEDAN LISTOS PARA REFERENCIARSE CUANDO SEAN NECESARIOS
    
    'POR EJEMPLO LOS PngBoton necesitan una imagen en inicarPNG que se puede acceder como
    'objFull.getFIMG.GetFileById("PNGUnSel").Path
    
    Exit Function

ErrLoadFull:
    terr.AppendLog "skin040", terr.ErrToTXT(Err)
End Function

'poder acceder a los archivos que usa la clase del manu (oManu) desde afuera
Public Function getFIMG() As clsFullImgs
    Set getFIMG = FImg
End Function

'darle a oSimple las preopiedades que tiene oManu
'las basicas y las extras (propis)
Public Function UpdatePropiedades()
    
    'QUEDEAQUI
    'LISTA DE LOS QUE YA TIENE ESTA PROPIEDAD
    
    'OK Tipo = en_clsLabel
    'NO Tipo = en_clsDiscoManager
    'NO Tipo = en_clsMANU
    'NO Tipo = en_clsNumeriks
    'OK Tipo = en_clsPNGBoton
    'OK Tipo = en_clsTemasManager
    'OK Tipo = en_tbrPromociones2
    'OK Tipo = en_tbrTextoSelect
    
    'a medida que se agregue el metodo cargar propiedades agregarlo aqui
    terr.Anotar "skin052"
    oManu.CargarPropiedades oSimple
    
End Function
