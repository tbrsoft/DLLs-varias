VERSION 5.00
Begin VB.UserControl tbrMAP 
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ScaleHeight     =   2535
   ScaleWidth      =   3750
   Begin VB.PictureBox Contenedor 
      Height          =   1605
      Left            =   180
      ScaleHeight     =   1545
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   150
      Width           =   3015
      Begin VB.PictureBox MAPA 
         AutoSize        =   -1  'True
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   1935
         TabIndex        =   3
         Top             =   30
         Width           =   2000
      End
   End
   Begin VB.VScrollBar vSC 
      Height          =   1605
      Left            =   3210
      TabIndex        =   1
      Top             =   150
      Width           =   345
   End
   Begin VB.HScrollBar hSC 
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   1770
      Width           =   2985
   End
End
Attribute VB_Name = "tbrMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ACTUALmapa As String

Private CR(100, 100) As Long 'hasta 100 corrdenadas se pueden guardar
Private CRName(100) As String 'los nombre de cada una de ellas

Dim JS As New tbrJUSE.clsJUSE

'los archivo tmap estan conformados por un JPG y un archivo de texto con las ubicaciones
'de cada una de las corrdenadas de referencia

'primero que todo necesito un JPG o BMP o GIF como base del mapa
Public Function AddBaseMap(xArch As String) As Long
    
End Function

'despues tenemos una AB (no tiene modificacion) de las corrdenadas de referencia

Public Function AddCoordenada(X As Long, Y As Long, NameCoord As String) As Long
    'me entrega las coordenadas en un zoom 1:1
    
    'ver que no exista la misma definicion
    
    Dim H As Long, EXIS As Boolean
    EXIS = False
    For H = 1 To 100
        If LCase(CRName(H)) = LCase(NameCoord) Then
            EXIS = True
            Exit For
        End If
    Next H
    
    If EXIS Then
        AddCoordenada = -1
        Exit Function
    End If
    
    'buscar un hueco en la matriz
    For H = 1 To 100
        If CRName(H) = "" Then
            CRName(H) = NameCoord
            CR(H, H) = 0
            AddCoordenada = H
            Exit Function
        End If
    Next H
End Function

Public Function RemoveCoordenada(NameCoord As String) As Long
    For H = 1 To 100
        If LCase(CRName(H)) = LCase(NameCoord) Then
            CRName(H) = ""
            Exit For
        End If
    Next H
    
    RemoveCoordenada = 0
End Function

Public Function ReleaseCoordenadas() As Long
    Erase CR
    Erase CRName
End Function

'despues ya se puede grabar como tMap
Public Function SaveTMap() As Long
    'verificar que la imagen sea valida
    'y que haya al menos un punto

End Function

'tambien se puede abrir uno preexistente
Public Function LoadTMap() As Long

End Function

Public Sub CargarubicMAPA()
    'hay archivos .ias que son coordenadas
    Dim ARCH As String
    lstPuntos.Clear
    Dim TE As TextStream
    Dim mapaARCH As String
    ARCH = Dir(AP + "ubic\*.ias")
    Do While ARCH <> ""
        Set TE = FSO.OpenTextFile(AP + "ubic\" + ARCH, ForReading, False)
        mapaARCH = TE.ReadLine
        'mostrar solo los que correspondan con el mapa mostrado
        If mapaARCH = ACTUALmapa Then lstPuntos.AddItem FSO.GetBaseName(ARCH)
        ARCH = Dir
    Loop
End Sub

Public Sub CargarMAPA(ARCH As String)
    MAPA.AutoRedraw = True
    MAPA.Picture = LoadPicture(ARCH)
    MAPA.Top = 0
    MAPA.Left = 0
    hSC.Min = 0
    hSC.Max = MAPA.Width - Contenedor.Width
    vSC.Min = 0
    vSC.Max = MAPA.Height - Contenedor.Height
    hSC.LargeChange = MAPA.Width / 20
    hSC.SmallChange = MAPA.Width / 200
    vSC.LargeChange = MAPA.Height / 20
    vSC.SmallChange = MAPA.Height / 200
    hSC = -MAPA.Left
    vSC = -MAPA.Top
    ACTUALmapa = ARCH
End Sub

Private Sub UserControl_Initialize()
    Me.AutoRedraw = True
    CargarMAPA FlMapas + "\villa.jpg"
    CargarubicMAPA
End Sub

Private Sub hSC_Change()
    MAPA.Left = -hSC
End Sub

Public Function ShowCoordenada() As Long
    'si hay alguno elegido
    If lstPuntos.SelCount > 0 Then
        'mostrar la parte del mapa que corresponde
        Dim LF As String, TP As String, MAP As String
        Dim TE As TextStream
        Set TE = FSO.OpenTextFile(AP + "ubic\" + lstPuntos + ".ias", ForReading, False)
        MAP = TE.ReadLine
        LF = TE.ReadLine
        TP = TE.ReadLine
        TE.Close
        MAPA.Left = Val(LF)
        MAPA.Top = Val(TP)
        'reubicar hsc y vsc (barras)
        hSC = -MAPA.Left
        vSC = -MAPA.Top
    End If
End Function

Private Sub vSC_Change()
    MAPA.Top = -vSC
End Sub

Public Function SelMapa() As String
    CmDlg.CancelError = False
    CmDlg.InitDir = FlMapas
    CmDlg.Filter = "Mapas de tbrSoft (*.tmap)|*.tmap"
    CmDlg.ShowOpen
    If CmDlg.FileName <> "" Then
        Dim ArchSel As String
        ArchSel = CmDlg.FileName
        CargarMAPA ArchSel
        CargarubicMAPA
    End If
End Function
