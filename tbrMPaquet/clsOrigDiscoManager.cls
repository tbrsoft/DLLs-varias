VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrigDiscoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'paquete de origenes de discos
'guarda tambien la lista de unidades que se van a usar con la finalidad de permitir copias o no segun espacio disponible

Dim ORs() As New clsOrigDisco  'paquete que maneja
Public Type UNI
    DrvLetter As String
    MbPedidos As Long 'cantidad de MB que se desean agregar
End Type

Dim UNIS() As UNI 'cada una de las letras de unidad que usan todos los origenes
Dim FSo As New Scripting.FileSystemObject

Public Function AddOrig(mPath As String) As clsOrigDisco
    
    Dim k As Long
    k = UBound(ORs) + 1
    ReDim Preserve ORs(k)
    ORs(k).Path = mPath
    Set AddOrig = ORs(k)
    
    'ver si es una nueva unidad ademas
    Dim m As Long
    Dim newUni As String 'vale algo si hay una nueva
    newUni = ORs(k).GetDriveLetter
    For m = 1 To UBound(UNIS) 'ver si es una ueva unidad
        If LCase(UNIS(m).DrvLetter) = LCase(newUni) Then
            newUni = "" 'lo vacio para que no se agregue
            ORs(k).iUni = m 'dejar referencia a que unidad esta para controlar peso de copias y discos llenados
        End If
    Next m
    
    If newUni <> "" Then
        m = UBound(UNIS) + 1
        ReDim Preserve UNIS(m)
        UNIS(m).DrvLetter = newUni
        ORs(k).iUni = m 'dejar referencia a que unidad esta para controlar peso de copias y discos llenados
    End If
    
End Function

'es muy generico pero indica que alguna de las unidades se va a llenar con el copiado propuesto
'devuelve LA LETRA DE LA UNIDAD que se va a llenar
Public Function HayEspacioOK() As String
    'indica si todo lo que se desea copiar no llenara ningun disco duro
    Dim k As Long
    For k = 1 To UBound(ORs)
        UNIS(ORs(k).iUni).MbPedidos = UNIS(ORs(k).iUni).MbPedidos + ORs(k).GetMBaCopiar
    Next k
    
    'verficar cada uno
    Dim DR As Drive
    For k = 1 To UBound(UNIS)
        Set DR = FSo.GetDrive(UNIS(k).DrvLetter)
        Dim FreeMB As Long
        FreeMB = (((DR.FreeSpace / 1024) / 1024) * 0.9) 'SOLO HASTA EL 90%
        If UNIS(k).MbPedidos > FreeMB Then
            HayEspacioOK = UNIS(k).DrvLetter
            Exit Function
        End If
    Next k
    
    HayEspacioOK = ""
    
End Function

Private Sub Class_Initialize()
    ReDim ORs(0)
    ReDim UNIS(0)
End Sub

Public Function GetOrig(mPath As String) As clsOrigDisco
    'devuelve un origen segun el path, ya debe estar cargado
    Dim k As Long
    For k = 1 To UBound(ORs)
        If LCase(mPath) = LCase(ORs(k).Path) Then
            Set GetOrig = ORs(k)
            Exit Function
        End If
    Next k
    
    Set GetOrig = ORs(0) 'vacio, solo llega aqui si no esta
End Function

Public Function GetOrigByIndex(i As Long) As clsOrigDisco
    Set GetOrigByIndex = ORs(i)
End Function

'devuelve todos los paths separados por algun caracter
Public Function GetFullOrigString(separator As String) As String
    Dim k As Long, t As String
    t = ""
    For k = 1 To UBound(ORs)
        t = t + ORs(k).Path
        If k < UBound(ORs) Then t = t + separator
    Next k
    
    GetFullOrigString = t
End Function

Public Function GetTotalOrs() As Long
    GetTotalOrs = UBound(ORs)
End Function
