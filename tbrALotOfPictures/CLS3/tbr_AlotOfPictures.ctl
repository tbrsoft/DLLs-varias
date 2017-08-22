VERSION 5.00
Begin VB.UserControl tbr_AlotOfPictures 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image IM 
      Height          =   795
      Index           =   0
      Left            =   2580
      Top             =   1170
      Width           =   1065
   End
End
Attribute VB_Name = "tbr_AlotOfPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'dll para almacenar y distribuir imagenes

'arranco pensando que las imagenes esten en otro espacio de memoria que no sea del _
    ejecutable para aliviarlo
'sin saber si el traspaso de uno a otro sea lento

Dim sKey() As String 'identificador unico de cada imagen
'sera su path ya que no se repetira espero

Dim FSO As New Scripting.FileSystemObject

Public Function AddImage(kKey As String, LoadToo As Boolean) As Long
    'si ya esta no lo agrego
    AddImage = 1 'valor si ya estaba
    
    If GetId(kKey) = 0 Then
        Dim H As Long
        H = UBound(sKey) + 1
        ReDim Preserve sKey(H)
        sKey(H) = kKey
        If LoadToo Then
            On Local Error GoTo FallaImg
            Load IM(H)
            IM(H).Picture = LoadPicture(kKey)
            IM(H).Tag = kKey
        End If
        AddImage = 0
    End If
    Exit Function
    
FallaImg:
    'eliminarla!
    H = UBound(sKey) - 1
    ReDim Preserve sKey(H)
    AddImage = -1
End Function

Public Function ClearAll()
    ReDim sKey(0)
    'ReDim SPic(0)
    ReDim MyBitMap(0)
End Function

Private Function GetId(kKey As String)
    'ver si existe
    GetId = 0 'bandera de que no existe
    Dim H As Long
    For H = 1 To UBound(sKey)
        If LCase(sKey(H)) = LCase(kKey) Then
            GetId = H 'ya existe
            Exit Function
        End If
    Next H
End Function

Public Function SetPicture(sPathF As String) As StdPicture
    If IM.Count > 1 Then
        Set SetPicture = IM(getI(sPathF)).Picture
    Else
        Set SetPicture = LoadPicture(sPathF)
    End If
End Function

Private Function getI(sFile As String)
    Dim H As Long
    For H = 1 To IM.Count - 1
        If LCase(sFile) = LCase(IM(H).Tag) Then
            getI = H
            Exit Function
        End If
    Next H
    getI = 0
End Function

Public Sub UnloadAll()
    Dim H As Long
    For H = 1 To IM.Count - 1
        IM(H).Picture = LoadPicture
        Unload IM(H)
    Next H
End Sub

Private Sub UserControl_Initialize()
    ClearAll
End Sub
