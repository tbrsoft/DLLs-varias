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

'no se que otra forma puedo acceder al padre para volver a pintarlo cuando cambien sus propiedades
Private PDR As ObjFullPadre

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
Public Tipo As tipoManu 'tipo del objeto del manu
Private fIMG As New clsFullImgs

Private mPath As String 'ultimo lugar donde se grabo o leyo el archivo

'para que solo se abra una sola vez el GuiEdit uso un obeto por cada objFull
Dim F_edit As frmEditManu

'cargar oManu segun corresponda
Public Function CreateManu()
    terr.Anotar "skin029", Tipo

    Dim X As Object 'valor de retorno
    
    If Tipo = en_clsMANU Then Set X = New clsMANU
    If Tipo = en_clsLabel Then Set X = New clsLabel
    If Tipo = en_clsNumeriks Then Set X = New clsNumeriks
    If Tipo = en_clsPNGBoton Then Set X = New clsPNGBoton
    If Tipo = en_clsTemasManager Then Set X = New clsTemasManager
    If Tipo = en_tbrTextoSelect Then Set X = New tbrTextoSelect
    If Tipo = en_tbrPromociones2 Then Set X = New tbrPromociones2
    If Tipo = en_clsDiscoManager Then Set X = New clsDiscoManager
    
    Set oManu = X
    
    terr.Anotar "skin030"
End Function

'hacer un objeto FullImg que contenga todos los archivos que necesite esta clase
'de modo que despues se pueda volver con un load
Public Function Save(Optional sFile As String = "")

    On Local Error GoTo ErrSaveFull
    terr.Anotar "skin031", sFile
    
    Save = 1
    
    If sFile = "" Then 'quiere regrabar el mismo
        If fso.FileExists(mPath) = False Then
            terr.AppendLog "TrySaveOBJFail!!"
            Exit Function
        Else
            sFile = mPath
        End If
    End If
    
    'parte comun a todos los objetos

    'hacer un archivo de texto con toda la config
    'uso el tbrConfig

    Dim C5 As New tbrConfig05.clstbrCfg05
    C5.SetPathError ap + "cfg005_4.log"
    Dim fCfg As String
    fCfg = tmpFolder + "cfg_" + CStr(Tipo) + "_" + oSimple.SGOName + ".txt"
    'limpiarlo si existe
    If fso.FileExists(fCfg) Then fso.DeleteFile fCfg, True
    
    Save = 2
    terr.Anotar "skin032", fCfg
    
    C5.SetMainFile fCfg

    Save = 3
    'seguiraqui asegurarse sincronizacion entre estas propiedades (oSimple) y las oManu
    UpdatePropiedades 'cargo valores reales del objeto del manu en una estructura mas linda en un tbrSGObject
    
    terr.Anotar "skin033"
    Save = 4
    '******************************************
    'estos son calculados y asignados por el SGO
    C5.SetV "X", CStr(oSimple.X)
    C5.SetV "Y", CStr(oSimple.Y)
    C5.SetV "Ancho", CStr(oSimple.W)
    C5.SetV "Alto", CStr(oSimple.H)
    C5.SetV "Estirable", CStr(CLng(oSimple.Estirable))
    C5.SetV "AlignH", CStr(oSimple.AlignementH)
    C5.SetV "AlignV", CStr(oSimple.AlignementV)
    '******************************************
    
    'las propiedades diferentes de las clases del manu se agregan como "Propis" en oSimple
    'cada una tiene diferentes pero se cargan de un solo modo como colecciones de propis aqui para simplificar todo
    Save = 5
    Dim H As Long, P As clsPropis
    For H = 1 To oSimple.GetPropCantidad
        Set P = oSimple.GetPropByID(H)
        C5.SetV P.NameProp, P.ValueProp
        terr.Anotar "skin034", H, P.NameProp, P.ValueProp
    Next H
    Save = 6
    terr.Anotar "skin035"
    C5.GrabarConfig
    
    terr.Anotar "skin036"
    
    Save = 7
    'empaquetar todo
    Dim F As clsFullImgs_FILE
    Set F = fIMG.AddFileByPath(fCfg, , "BASE_CONFIG") 'config de texto
    F.sID = "BASE_CONFIG"
    
    terr.Anotar "skin037"
    'ver si la clase del manu tiene imagenes que guardar
    'pudiera ser posible que todos los objetos del manu tengan un FullImg para respetar una estructura comun
    'de todas formas todavia no parece necesario y lo manjeo segun caso
    Save = 8
    terr.Anotar "skin037-d", fIMG.GetFileCantidad
    fIMG.AddFilesFromOtherFIMG oManu.GetFIMGs
    
    terr.Anotar "skin038"
    Save = 9
    fIMG.SaveFile sFile
    
    terr.Anotar "skin039"
    mPath = sFile 'por si quiere grabar
    Save = 0
    Exit Function
    
ErrSaveFull:
    terr.AppendLog "skn3021", terr.ErrToTXT(Err)
    
End Function

'indica que archivo descomprimir y donde grabar los archivos extraidos
Public Function Load(sFile As String, sFolder As String) As Long

    On Local Error GoTo ErrLoadFull
    terr.Anotar "skin044", sFile, sFolder
    
    'descomprimir el archivo
    
    Dim H As Long
    H = fIMG.OpenSkinFile(sFile, sFolder)
    
    terr.Anotar "skin045", H
    If H > 0 Then 'error al abrirse archivo
        terr.AppendLog "NoOpenSk:" + CStr(H) + ":" + sFile + ":" + sFolder
        Exit Function
    End If
    
    'buscar entre todos los archivos cual es de la config (empieza con "CFG_")
    Dim FI As clsFullImgs_FILE
    Set FI = fIMG.GetFileByName("cfg_*") 'buscar un archivo que empiece con "cfg_" (SOLO DEBE HABER UNO)
    
    terr.Anotar "skin046"
    'ver de que tipo y que nombre tenia
    Dim n As String, SP() As String
    n = FI.BaseName
    SP = Split(n, "_")
    
    terr.Anotar "skin047", n, SP(1), SP(2)
    '***********************************
    Tipo = CLng(SP(1))          'ASIGNAR estos valores importantes
    oSimple.SGOName = SP(2)
    '***********************************
    'crear el objeto del manu (solo segun el tipo)
    CreateManu
    
    'abrir la config y aplicar los valores
    Dim fCfg As String 'archivo de la config
    fCfg = FI.path
    
    terr.Anotar "skin048"
    Dim C5 As New tbrConfig05.clstbrCfg05
    C5.SetPathError ap + "cfg005_5.log"
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
                oSimple.X = C5.GetLong("X")
            Case "y"
                oSimple.Y = C5.GetLong("Y")
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
                oSimple.SetProp CF, C5.GetV(CF)
        
        End Select
        
        '******************************************
        
    Next J
    
    terr.Anotar "skin051"
    '-------------MONTAR ARCHIVOS EXTERNOS SI HAY -------------------------------
    'LAS IMAGENES O ARCHIVOS EXTERNOS QUEDAN LISTOS PARA REFERENCIARSE CUANDO SEAN NECESARIOS
    'con objFull.getFIMG.get...
    
    
    'POR EJEMPLO LOS PngBoton necesitan una imagen en inicarPNG que se puede acceder como
    'objFull.getFIMG.GetFileById("PNGUnSel").Path
    '-----------------------------------------------------------------------------
    
    mPath = sFile 'por si quiere grabar sin cambiar el nombre o sea hizo el load para editarlo
    Load = 0
    Exit Function

ErrLoadFull:
    Load = Err.Number
    terr.AppendLog "skin040", terr.ErrToTXT(Err)
    
End Function

'poder acceder a los archivos que usa la clase del manu (oManu) desde afuera
Public Function getFIMG() As clsFullImgs
    Set getFIMG = fIMG
End Function

'darle a oSimple las preopiedades que tiene oManu [PARA ESTO ES NECESARIO QUE OMANU HAYA SIDO INICIALIZADO]
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

Private Sub Class_Initialize()
    Set oSimple = New tbrSuperGuiObject
End Sub

'editar sus propiedades, necesito saber el padre, no se de que otra forma puedo acceder a el
Public Sub GuiEdit(P As ObjFullPadre)
    Set PDR = P
    If F_edit Is Nothing Then
        Set F_edit = New frmEditManu
    End If
    F_edit.SetObjFull Me
    F_edit.Caption = "Obj: " + Me.oSimple.SGOName + " / Tipo: " + CStr(Me.Tipo)
    F_edit.Show
End Sub

'lo uso para redibujar cuando hago cambios en el formulario de edicion y nada mas
Public Function GetPadre() As ObjFullPadre
    Set GetPadre = PDR 'seguiraqui deberia yo marcar error si se pide esto y el padre no esta seteado (solo, se hace en guiEdit)
End Function
