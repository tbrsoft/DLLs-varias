VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cosas vinculadas a la interfase gráfica

'independiente del manager de objetos para que aqui se definan cosas propias de cada programa que lo usa
'como el load, el save y los pedidos desde el sistema

'aqui defino tambien los objetos a usar realmente para simplificar las llamadas a cada uno
Private SG As tbrSuperGUI  'madre de todos los objetos
Private mPath As String 'archivo que estoy manejando

Public Property Get path() As String
    path = mPath
End Property

Public Function MNG() As tbrSuperGUI 'manager de los objetos
    Set MNG = SG
End Function

Public Sub SetPathLog(pathLog As String)
    terr.Anotar "skin055"
    
    terr.FileLog = pathLog 'seguiraqui asegurarse que se asigne
    terr.LargoAcumula = 901
    
End Sub

Public Function SaveFile(sFile As String)
    On Local Error GoTo ERRsvGR
    terr.Anotar "skin060", sFile
        
    If sFile = "" Then 'quiere grabar el que estaba
        sFile = mPath
    Else
        mPath = sFile
    End If
    
    'se grabara como temporal porque es posible que falle esta grabacaion y quede dañado el archivo!!!
    sFile = mPath + ".temp"
    If fso.FileExists(sFile) Then fso.DeleteFile sFile, True
    
    SaveFile = 1
    'pasar por cada objeto y grabarlo
    'en un juse
    Dim Ju As New tbrjuse2.clsJUSE
    Ju.Archivo = sFile
    
    'grabar todos los padres
    Dim tmpPadre As ObjFullPadre
    Dim M As Long
    Dim cadaFile As String
    SaveFile = 2
    
    For M = 1 To SG.GetPadresMaxID
        'cada uno se inicializa diferente
        
        Set tmpPadre = SG.GetPadresByID(M)
        
        cadaFile = sFile + CStr(M) + ".padre" 'para saber como abrirlo despues
        
        terr.Anotar "skin061c", tmpPadre.sName, cadaFile
        tmpPadre.Save cadaFile
        
        Ju.AddFile cadaFile
        
    Next M
    
    SaveFile = 4
    
    '******************************************'******************************************
    '******************************************'******************************************
    
    'CONFIGURACION GENERAL DE ESTA INTERFASE GRAFICA
    Dim C5 As New tbrConfig05.clstbrCfg05
    C5.SetPathError ap + "cfg005.log"
    
    Dim fCfg As String
    fCfg = tmpFolder + "CONFIG_GUI.txt"
    'limpiarlo si existe
    If fso.FileExists(fCfg) Then fso.DeleteFile fCfg, True
    
    Save = 41
    terr.Anotar "skin032-f", fCfg
    
    C5.SetMainFile fCfg

    Save = 42
    
    terr.Anotar "skin033-g"
    Save = 4
    
    '******************************************
    'estos son calculados y asignados por el SGO
    C5.SetV "HechoParaPixlesAncho", CStr(HechoParaPixlesAncho)
    C5.SetV "HechoParaPixlesAlto", CStr(HechoParaPixlesAlto)
    '******************************************
    Save = 43
    terr.Anotar "skin033-h"
    C5.GrabarConfig
    
    Save = 44
    Ju.AddFile fCfg
    terr.Anotar "skin033-i"
    Save = 45
    '******************************************'******************************************
    '******************************************'******************************************
    
    Ju.Unir
    SaveFile = 5
    'borrar ahora todos los padres y los hijos que quedaron sueltos!!
    'ya estan dentro del juse !
    For M = 1 To SG.GetPadresMaxID
        cadaFile = sFile + CStr(M) + ".padre" 'para saber como abrirlo despues
        terr.Anotar "skin061t", cadaFile
        fso.DeleteFile cadaFile, True
    Next M
    
    SaveFile = 7
    'borrarlo si existia, ya hay una nueva version temporal lista
    If fso.FileExists(mPath) Then fso.DeleteFile mPath, True
    'el temporal se grabo ok, ya puedo reemplazar el ofical
    fso.MoveFile mPath + ".temp", mPath
    
    SaveFile = 0 'sale ok
    Exit Function
    
ERRsvGR:
    terr.AppendLog "ERRsvGR", terr.ErrToTXT(Err)
    
End Function

'cargar config
Public Function LoadFile(sFile As String, DestFolder As String) As Long
    
    On Local Error GoTo errSG
    
    terr.Anotar "skin070", sFile
    mPath = sFile
    
    LoadFile = 1
    'pasar por cada objeto y grabarlo
    'en un juse
    Dim Ju As New tbrjuse2.clsJUSE
    
    Dim res As Long
    res = Ju.ReadFile(sFile)
    
    terr.Anotar "skin071", res
    If res = 1 Then
        terr.AppendLog "JusBroke072"
        LoadFile = 2
        Exit Function
    End If
    
    terr.Anotar "skin073", Ju.CantArchs

    LoadFile = 3
    Dim Cada As String, A2 As Long
    
    For A2 = 1 To Ju.CantArchs

        Cada = Ju.GetListFiles(A2, False)
        terr.Anotar "skin074", A2, Cada

        If Cada = "" Then
            terr.AppendLog "FailJUS075"
            LoadFile = 4
            Exit Function
        End If

        Ju.Extract DestFolder, A2
        
        Dim esteArch As String, esteRef As String
        esteArch = DestFolder + Cada
        
        'abrir los padres
        If LCase(fso.GetExtensionName(esteArch)) = "padre" Then
        
            'necesito una carpeta para cada archivo [que sea un padre]
            Dim FolExtract As String
            FolExtract = DestFolder + CStr(A2)
            terr.Anotar "skin074g", FolExtract
            If fso.FolderExists(FolExtract) Then
                fso.DeleteFolder FolExtract, True
            End If
            
            terr.Anotar "skin074h", Cada
            fso.CreateFolder FolExtract
        
            Dim tmpPadre As ObjFullPadre
            Set tmpPadre = SG.AddPadre("temporal")
            tmpPadre.Load esteArch, FolExtract
            
        End If
        
        'CONFIG_GUI.txt
        If esteArch = "CONFIG_GUI.txt" Then
            LoadCFGInterna esteArch
        End If
        
    Next A2

    LoadFile = 0 'sale ok
    
    
    Exit Function
errSG:
    LoadFile = 2
    terr.AppendLog "errSG", terr.ErrToTXT(Err)
    
End Function

Private Sub LoadCFGInterna(fCfg As String)
    
    terr.Anotar "skin048-j"
    Dim C5 As New tbrConfig05.clstbrCfg05
    C5.SetPathError ap + "cfg005_2.log"
    C5.AbrirCfg fCfg
    
    HechoParaPixlesAncho = C5.GetLong("HechoParaPixlesAncho")
    HechoParaPixlesAlto = C5.GetLong("HechoParaPixlesAlto")
    
    'asegurarse que no este en cero !!!
    If HechoParaPixlesAncho = 0 Then HechoParaPixlesAncho = 1024
    If HechoParaPixlesAlto = 0 Then HechoParaPixlesAlto = 768
    
    terr.Anotar "skin049-k", h1, H2

End Sub

'cerrar todos los graficos y vaciar todos los elementos
Public Sub ClearAll()
    'pasar por cada uno de los padres y cerrarle todos los graficos a sus hijos
    Dim M As Long
    For M = 1 To MNG.GetPadresMaxID
        MNG.GetPadresByID(M).INIT_GRAPH "CLOSE" 'cierra graficos iterando por cada objeto
    Next M
    
    MNG.Clear 'elimina los objetos (redim a cero)
End Sub

Private Sub Class_Initialize()
        
    'el manager simepre listo
    Set SG = New tbrSuperGUI
    
    'MUY IMPORTANTE ESTE FRM LO USAN TODOS LOS OBJETOS DEL MALDITO MANU
    Load F_TEMP
    Set qAlgunFormulario = F_TEMP
    
    'definr una carpeta temporal
    tmpFolder = App.path
    If Right(tmpFolder, 1) <> "\" Then tmpFolder = tmpFolder + "\"
    tmpFolder = tmpFolder + "tmp\"
    
    If fso.FolderExists(tmpFolder) = False Then
        fso.CreateFolder tmpFolder
    End If
    
End Sub

Public Function SetParaPantallaPixeles(wi As Long, he As Long)
    HechoParaPixlesAncho = wi
    HechoParaPixlesAlto = he
End Function
