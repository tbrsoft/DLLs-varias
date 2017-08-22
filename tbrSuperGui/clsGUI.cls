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

Private SG As tbrSuperGUI 'madre de todos los objetos

Public Function MNG() As tbrSuperGUI 'manager de los objetos
    Set MNG = SG
End Function

'cosas obligatorias
Public Sub INIT(pxAncho As Long, pxAlto As Long)
    Set SG = New tbrSuperGUI
    
    SG.HechoParaPixlesAncho = pxAncho '1024 - 800
    SG.HechoParaPixlesAlto = pxAlto   ' 768 - 600
End Sub

'seguiraqui el archivo sgo.txt por ahora contiene todas estas configs
Public Function LoadFile(sFile As String) As Long
    
    On Local Error GoTo errSG
    tERR.Anotar "fafa", sFile
    If fso.FileExists(sFile) = False Then
        tERR.AppendSinHist "fafb:" + sFile
        LoadFile = 1
        Exit Function
    End If
    
    Dim TE As TextStream
    
    tERR.Anotar "fafc"
    Dim lastFRM As String 'como las propiedades estan sueltas siempre se asignan al ultimo frm & name visto
    Dim lastName As String
    Dim lastOB As ObjFULL  'para ir poniendole las propiedades que aparezcan
    
    Dim SP1() As String 'objetos a leer
    Dim SP2() As String 'valores dentro de cada cosa leida
    Set TE = fso.OpenTextFile(sFile, ForReading)
        re = TE.ReadAll
    TE.Close
    
    tERR.Anotar "fafd"
    
    SP1 = Split(re, ";")
    Dim k As Long
    For k = 0 To UBound(SP1)
        tERR.Anotar "fafe", k, SP1(k)
        'limpiarla de espacios y mayusculas
        SP1(k) = Trim(LCase(SP1(k))) 'si hubiera muchos espacios o tabulaciones al inicio las quita
        
        'cortar en partes separadas por espacios cada propiedad
        SP2 = Split(SP1(k))
        
        'sacarle saltos de carro y cosas chongas
        SP2(0) = Replace(SP2(0), vbCr, "")
        SP2(0) = Replace(SP2(0), vbLf, "")
        SP2(0) = Replace(SP2(0), vbCrLf, "")
        SP2(0) = Replace(SP2(0), Chr(9), "") 'tabulacion
        
        tERR.Anotar "fafg", SP2(0)
        
        Select Case SP2(0)
            Case "frm"
                lastFRM = SP2(1)
        
            Case "name"
                tERR.Anotar "fafh", SP2(1)
                lastName = SP2(1) 'aqui define que se creara un objeto
                lastOB = SG.AddSgo(lastFRM + "_" + lastName)   'nombre compuesto para que sea unico
    
            Case "rect"
                lastOB.oSimple.x = CLng(SP2(1))
                lastOB.oSimple.Y = CLng(SP2(2))
                lastOB.oSimple.W = CLng(SP2(3))
                lastOB.oSimple.H = CLng(SP2(4))
        
            Case "align"
                lastOB.oSimple.AlignementH = CLng(SP2(1))
                lastOB.oSimple.AlignementV = CLng(SP2(2))
        
            Case "estirable"
                lastOB.oSimple.Estirable = CBool(CLng(SP2(1)))
        
            Case "tipo"
                tERR.Anotar "fafi", SP2(1)
                'SEGUIRAQUI
                'esta clase forma parte de la dll del manu y crea objetos segun sus clases aqui mismo
                'set lastOB.obj = new objetoManuDelTipoQueSea()
                'deberia guardar el indice o cosas similares
                lastOB.Tipo = SP2(1)
                Set lastOB.oManu = SG.GetNewManu(lastOB.Tipo) 'obtener un nuevo objeto del tipo de esta funcion que entrega nuevas instancias
            
            Case "propiedadesDelManu"
                'por cada propiedad le asigno a lastOB lo que el manu me diga
                
                
            Case Else
                'cometarios y otras cosas
            
        End Select
            
    Next k
    tERR.Anotar "fafj"
    '//////////////////////////////////////////////
    'acomoda todos los x2 finales segun estirable y alineacion
    SG.TraslateALL ' estirar o no segun corresponda
    '//////////////////////////////////////////////
    
    tERR.Anotar "fafk"
    LoadFile = 0
    
    Exit Function
errSG:
    LoadFile = 2
    tERR.AppendLog "errSG", tERR.ErrToTXT(Err)
    
End Function
