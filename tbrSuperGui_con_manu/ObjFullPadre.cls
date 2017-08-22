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

Private FImg As New clsFullImgs 'para la imagen de fondo y el archivo de cfg y otras cosas que puedan surgir en el futuro

Public Sub Save(sFile As String)
    On Local Error GoTo ErrSaveFull
    terr.Anotar "skin080", sFile, sName

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
    '******************************************
    
    C5.GrabarConfig
    
    terr.Anotar "skin082"
    
    'empaquetar todo
    Dim F As clsFullImgs_FILE
    Set F = FImg.AddFileByPath(fCfg) 'config de texto
    F.sID = "BASE_CONFIG_PADRE"
    
    terr.Anotar "skin083"
    'ver si usa imagen de fondo
    If fso.FileExists(pathImgFondo) Then
        Dim F2 As clsFullImgs_FILE
        Set F2 = FImg.AddFileByPath(pathImgFondo)
        F2.sID = "FondoPadre"
    End If
    
    terr.Anotar "skin084"
    
    FImg.SaveFile sFile
    
    terr.Anotar "skin086"
    Exit Sub
    
ErrSaveFull:
    terr.AppendLog "skn3085", terr.ErrToTXT(Err)
End Sub

Public Sub Load(sFile As String, sFolder As String)
    On Local Error GoTo ErrLoadFull
    terr.Anotar "skin087", sFile, sFolder
    
    'descomprimir el archivo
    
    Dim H As Long
    H = FImg.OpenSkinFile(sFile, sFolder)
    
    terr.Anotar "skin088", H
    If H > 0 Then 'error al abrirse archivo
        terr.AppendLog "No089OpenSk:" + CStr(H) + ":" + sFile + ":" + sFolder
        Exit Sub
    End If
    
    'buscar entre todos los archivos cual es de la config (empieza con "CFG_")
    Dim FI As clsFullImgs_FILE
    Set FI = FImg.GetFileByName("cfg_padre")
    
    terr.Anotar "skin090"
    
    'abrir la config y aplicar los valores
    Dim fCfg As String 'archivo de la config
    fCfg = FI.path
    
    terr.Anotar "skin091"
    Dim C5 As New tbrConfig05.clstbrCfg05
    C5.AbrirCfg fCfg
    
    terr.Anotar "skin092", fCfg
    sName = C5.GetV("sName")
    
    terr.Anotar "skin093", sName
    
    'si no existe queda en "", no es problema (registrara un AppendLog pero despues lo velo) seguiraqui
    pathImgFondo = FImg.GetFileBysID("FondoPadre").path
    
    Exit Sub
    
ErrLoadFull:
    terr.AppendLog "skin040", terr.ErrToTXT(Err)
End Sub
