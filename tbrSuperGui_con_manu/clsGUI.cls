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

Public Function MNG() As tbrSuperGUI 'manager de los objetos
    Set MNG = SG
End Function

'cosas obligatorias
Public Sub INIT(pxAncho As Long, pxAlto As Long, Optional pathLog As String = "")
    Set SG = New tbrSuperGUI
    
    SG.HechoParaPixlesAncho = pxAncho '1024 - 800
    SG.HechoParaPixlesAlto = pxAlto   ' 768 - 600
    
    If pathLog <> "" Then
        terr.FileLog = pathLog 'seguiraqui asegurarse que se asigne
    End If
    
    terr.LargoAcumula = 8001
    
End Sub

'seguiraqui el archivo sgo.txt por ahora contiene todas estas configs
Public Function LoadFile(sFile As String) As Long
    
    On Local Error GoTo errSG
    terr.Anotar "fafa", sFile
    If fso.FileExists(sFile) = False Then
        terr.AppendSinHist "fafb:" + sFile
        LoadFile = 1
        Exit Function
    End If
    
    Dim TE As TextStream
    
    terr.Anotar "fafc"
    Dim lastFRM As String 'como las propiedades estan sueltas siempre se asignan al ultimo frm & name visto
    Dim lastOB As objFULL  'para ir poniendole las propiedades que aparezcan
    
    Dim SP1() As String 'objetos a leer
    Dim SP2() As String 'valores dentro de cada cosa leida
    Set TE = fso.OpenTextFile(sFile, ForReading)
        re = TE.ReadAll
    TE.Close
    
    terr.Anotar "fafd"
    
    SP1 = Split(re, ";")
    Dim k As Long
    For k = 0 To UBound(SP1)
        terr.Anotar "fafe", k, SP1(k)
        'limpiarla de espacios y mayusculas
        SP1(k) = Trim(LCase(SP1(k))) 'si hubiera muchos espacios o tabulaciones al inicio las quita
        
        'cortar en partes separadas por espacios cada propiedad
        SP2 = Split(SP1(k))
        
        'sacarle saltos de carro y cosas chongas
        SP2(0) = Replace(SP2(0), vbCr, "")
        SP2(0) = Replace(SP2(0), vbLf, "")
        SP2(0) = Replace(SP2(0), vbCrLf, "")
        SP2(0) = Replace(SP2(0), Chr(9), "") 'tabulacion
        
        terr.Anotar "fafg", SP2(0)
        
        Select Case LCase(SP2(0))
            
            Case "exitfile" 'salir ha terminado
                Exit For
                
            Case "frm"
                lastFRM = SP2(1)
                SG.AddPadre lastFRM 'nuevo formulario o contenedor grafico
        
            Case "name"
                terr.Anotar "fafh", SP2(1)
                Set lastOB = SG.AddSgo(lastFRM, SP2(1))
    
            'propiedades preferenciales
            Case "rect"
                lastOB.oSimple.x = CLng(SP2(1))
                lastOB.oSimple.y = CLng(SP2(2))
                lastOB.oSimple.W = CLng(SP2(3))
                lastOB.oSimple.H = CLng(SP2(4))
        
            Case "align"
                lastOB.oSimple.AlignementH = CLng(SP2(1))
                lastOB.oSimple.AlignementV = CLng(SP2(2))
        
            Case "estirable"
                lastOB.oSimple.Estirable = CBool(CLng(SP2(1)))
        
            Case "tipo"
                terr.Anotar "fafi", SP2(1)
                'SEGUIRAQUI
                'esta clase forma parte de la dll del manu y crea objetos segun sus clases aqui mismo
                'set lastOB.obj = new objetoManuDelTipoQueSea()
                'deberia guardar el indice o cosas similares
                lastOB.Tipo = SP2(1)
                Set lastOB.oManu = SG.GetNewManu(lastOB.Tipo) 'obtener un nuevo objeto del tipo de esta funcion que entrega nuevas instancias
                
                If TypeOf lastOB.oManu Is clsTemasManager Then
                    TMP = TMP + "oManu = tm"
                End If
                
                
                'no lo inicializa aqui por que nos se han definido todas las propiedades y no se han traducido todas las coordenadas finales
                
            Case "prop" 'otras propiedades, no requieren existir en las clases
                'por cada propiedad le asigno a lastOB lo que el manu me diga
                'si el valor de la prop tuviera espacios seria un problema
                Dim realVal As String: realVal = ""
                Dim R As Long
                For R = 2 To UBound(SP2)
                    realVal = realVal + SP2(R)
                    If R < UBound(SP2) Then realVal = realVal + " "
                Next R
                
                lastOB.oSimple.AddProp Trim(SP2(1)), realVal
                
            Case Else
                'cometarios y otras cosas
            
        End Select
            
    Next k
    terr.Anotar "fafj"
    '//////////////////////////////////////////////
    'acomoda todos los x2 finales segun estirable y alineacion
    SG.TraslateALL ' estirar o no segun corresponda
    '//////////////////////////////////////////////

    'resta inciializar los graficos de todos
    
    terr.Anotar "fafk"
    LoadFile = 0
    
    Exit Function
errSG:
    LoadFile = 2
    terr.AppendLog "errSG", terr.ErrToTXT(Err)
    
End Function

'funcon para hacerle algo a TODOS los objetos graficos, inicializar y cerrra por ejemplo
Public Sub INIT_GRAPH(Optional sClose As String = "INIT")

    terr.Anotar "fafk-2", sClose
    On Local Error GoTo ERRGR

    'hacer todos los incializargraficos que correspondan segun indicaciones del manu!
    'se hace aqui por que ya se han traducido las coordenadas en el mismo load
    
    'seguiraqui - necesito los HDC de cada formulario
    'seguiraqui -
    
    Dim M As Long
    
    'todos usan el sgo de cada ID
    Dim tmpSGO As objFULL
    For M = 1 To SG.GetSgoMaxID
        'cada uno se inicializa diferente
        
        Set tmpSGO = SG.GetSgoByID(M)
        
        terr.Anotar "initfg_" + tmpSGO.Tipo + " : " + CStr(tmpSGO.oSimple.X2) + " : " + CStr(tmpSGO.oSimple.Y2) + " : " + CStr(tmpSGO.oSimple.W2) + " : " + CStr(tmpSGO.oSimple.H2)
        
        Select Case tmpSGO.Tipo
            
            Case "label" '**************LABEL DE CREDITOS,
                Dim OV_01 As clsLabel
                Set OV_01 = tmpSGO.oManu
                If sClose = "INIT" Then
                    OV_01.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColor"), _
                        tmpSGO.oSimple.GetProp_LNG("FntColorContorno")
                    
                    OV_01.IniciarGrafios tmpSGO.Padre.sHDC, _
                        tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, tmpSGO.oSimple.H2, _
                        tmpSGO.oSimple.GetProp_BOOL("alphaHabilitado"), tmpSGO.oSimple.GetProp_LNG("AlpahColorLong"), _
                        tmpSGO.oSimple.GetProp_LNG("AlphaCantidad")
                End If
                
                If sClose = "CLOSE" Then
                    OV_01.CerrarGraficos
                End If
            Case "lista" '*********LISTAS DE CANCIONES O DE RITMOS O DE LETRAS; ETC
            
                Dim OV_02 As clsTemasManager
                Set OV_02 = tmpSGO.oManu
                
                If sClose = "INIT" Then
                    OV_02.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColorUnSel"), _
                        tmpSGO.oSimple.GetProp_LNG("FontColorSel"), tmpSGO.oSimple.GetProp_LNG("FntColorContorno")
                    
                    OV_02.IniciarGrafios tmpSGO.Padre.sHDC, _
                        tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, tmpSGO.oSimple.H2, _
                        tmpSGO.oSimple.GetProp_BOOL("AlphaB"), tmpSGO.oSimple.GetProp_LNG("ColorSel")
                End If
                
                If sClose = "CLOSE" Then
                    OV_02.CerrarGraficos
                End If
                
            Case "listah" '*************listas horizontales ritmos, letras
            
                Dim OV_03 As tbrTextoSelect
                Set OV_03 = tmpSGO.oManu
                
                If sClose = "INIT" Then
                
                    OV_03.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColorUnSel"), _
                        tmpSGO.oSimple.GetProp_LNG("FontColorSel"), tmpSGO.oSimple.GetProp_LNG("FntColorContorno")
                    
                    OV_03.IniciarGrafios tmpSGO.Padre.sHDC, _
                        tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, tmpSGO.oSimple.H2
                
                End If
                
                If sClose = "CLOSE" Then
                    OV_03.CerrarGraficos
                End If
                
            Case "promociones2"
                Dim OV_04 As tbrPromociones2
                Set OV_04 = tmpSGO.oManu
                                
                If sClose = "INIT" Then
                
                    OV_04.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColor")
                        
                    OV_04.IniciarGrafios tmpSGO.Padre.sHDC, _
                        tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, tmpSGO.oSimple.W2, tmpSGO.oSimple.H2, 3, 40 'SEGUIRAQUI los tiempos de cambio de textos podrian ser configurables, estan fijos
                End If
                
                If sClose = "CLOSE" Then
                    OV_04.CerrarGraficos
                End If
            
            Case "discomanager"
                Dim OV_05 As clsDiscoManager
                Set OV_05 = tmpSGO.oManu
                OV_05.IniciarGrafios tmpSGO.Padre.sHDC, tmpSGO.oSimple.X2, tmpSGO.oSimple.Y2, _
                tmpSGO.oSimple.W2, tmpSGO.oSimple.H2, False, _
                OV_05.MostrarModoNumpad, tmpSGO.oSimple.GetProp_LNG("ColorDiscoSel")
    
                terr.Anotar "aabt30"
                
                If sClose = "INIT" Then
                    OV_05.IniciarFuente tmpSGO.oSimple.GetProp_STR("FontName"), _
                        tmpSGO.oSimple.GetProp_LNG("FontSize"), tmpSGO.oSimple.GetProp_BOOL("FontBold"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontItalic"), tmpSGO.oSimple.GetProp_BOOL("FontStrike"), _
                        tmpSGO.oSimple.GetProp_BOOL("FontUnderLine"), tmpSGO.oSimple.GetProp_LNG("FontColor"), _
                        tmpSGO.oSimple.GetProp_LNG("FntColorContorno")
                End If
            
                If sClose = "CLOSE" Then
                    OV_05.CerrarGraficos
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

Private Sub Class_Initialize()
    'MUY IMPORTANTE ESTE FRM LO USAN TODOS LOS OBJETOS DEL MALDITO MANU
    Load F_TEMP
    Set qAlgunFormulario = F_TEMP
End Sub
