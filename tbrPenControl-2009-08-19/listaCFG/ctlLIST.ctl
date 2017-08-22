VERSION 5.00
Begin VB.UserControl ctlLIST 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton btOk 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   2955
      Width           =   1000
   End
   Begin VB.CommandButton btCa 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   1215
      TabIndex        =   1
      Top             =   2955
      Width           =   1000
   End
   Begin VB.TextBox lbTITULO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   4005
   End
End
Attribute VB_Name = "ctlLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Clases necesarias para la lista:
'cada elemento de la lista
'Dim ListaListaElem As New clsTemaList
'el manager de los elementos
Dim ListaLista As New clsTemasManager

Private Const SRCCOPY = &HCC0020  ' used to determine how a blit will turn out
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'----------------------------------------------

Private mMargen As Long 'margen arriba abajo der izq
Private mMargenRenglones As Long

Private mForeColorNormal As Long
Private mForeColorSel As Long
Private mBackColorSel As Long
Private mBackColorFondo As Long

Private manageList As New clsVERListaSimple

Public Event Change(NewSel As String)
Public Event ClickOK()
Public Event ClickCancel()

Private Sub btCa_Click()
    RaiseEvent ClickCancel
End Sub

Private Sub btOk_Click()
    RaiseEvent ClickOK
End Sub

Public Function setManager(mng As clsVERListaSimple)
    Set manageList = mng
End Function

Public Function getManager() As clsVERListaSimple
    Set getManager = manageList
End Function

Public Function LoadList()
        
    ListaLista.ResetElementos
    ListaLista.IniciarFuente NegradaFrmManu, "Verdana", 8, True, True, False, False, vbWhite, vbWhite, RGB(80, 80, 80)
    ListaLista.IniciarGrafios UserControl.HDC, 3, 3, (UserControl.Width / 15) - 6, (UserControl.Height / 15) - (btOk.Height / 15) - 6, False, vbWhite
    
    terr.Anotar "qdj"
    
    Dim a As Long
    For a = 0 To UBound(manageList.GetStringListVisible)
        terr.Anotar "qdk", a, manageList.GetOpVisible(a)
        
        Dim L As New clsTemaList
        Set L = ListaLista.AgregarElemento
        L.Numero = -1
        L.Titulo = manageList.GetOpVisible(a)
        L.TagMisterioso = manageList.GetOpTag(a)
        
    Next a
    
    ListaLista.IniciarTouchScreen
    
    ReAcomodar
    terr.Anotar "qdl"
    UpdateSel 'marca el elegido
    terr.Anotar "qdm"
End Function

Public Sub SelNext()
    terr.Anotar "qdn", ListaLista.GetIndiceElegido, ListaLista.GetElegido.Titulo
    ListaLista.SelNext True
    terr.Anotar "qdo"
    UpdateSel 'mostrar el elegido
End Sub

Public Sub SelPrev()
    terr.Anotar "qdp", ListaLista.GetIndiceElegido, ListaLista.GetElegido.Titulo
    ListaLista.SelPrevious True
    terr.Anotar "qdq"
    UpdateSel 'mostrar el elegido
    terr.Anotar "qdr"
End Sub

Public Sub SelElegida() 'marcar como elegida la que corresponde (lo hago al iniciar uin combo buscando a partir de a cfg elegida en esa opcion)
    Dim a As Long
    'seguiraqui, asegurarse que elije bien de un valor grabado previamente
    
    For a = 1 To ListaLista.GetElementoCount
        terr.Anotar "qds", a, manageList.GetSelectOp, manageList.GetSelectOpInternal, manageList.GetSelectOpTMP
        
        Dim CompareTo As String 'si no hay nada elegido seguro hay algo temporal y eso debe mostrarse !
        If manageList.GetSelectOp = "NULL" Then
            CompareTo = manageList.GetSelectOpTMP
        Else
            CompareTo = manageList.GetSelectOp
        End If
        
        If ListaLista.GetElem(a).Titulo = CompareTo Then
            ListaLista.DefineElegidoByIndex a
            Exit For
        End If
    Next a
    terr.Anotar "qdt"
    UpdateSel
    terr.Anotar "qdu"
End Sub

Private Sub UpdateSel()
    terr.Anotar "qdx"
    manageList.DefineSelectFromID ListaLista.GetIndiceElegido
    If Not (ListaLista.GetElegido Is Nothing) Then
        RaiseEvent Change(ListaLista.GetElegido.Titulo)
    End If
    
    UserControl.Refresh
End Sub

Private Sub UserControl_Initialize()
    
    mMargen = 60
    mMargenRenglones = 60
    
    Me.Alignment = vbCenter
    
    mBackColorFondo = 0
    mBackColorSel = &HFF0000
    mForeColorNormal = vbWhite
    mForeColorSel = &HFFFFC0
    
    btOk.Font = "Verdana"
    btOk.FontSize = 8
    btOk.FontBold = False
    btCa.Font = "Verdana"
    btCa.FontSize = 8
    btCa.FontBold = False
    
End Sub

Private Sub UserControl_Resize()
    btOk.Top = UserControl.Height - btOk.Height - 30
    btOk.Left = 60
    btCa.Top = btOk.Top
    btCa.Left = btOk.Left + btOk.Width + 30
    
    ReAcomodar
End Sub

Public Property Get Alignment() As AlignmentConstants
    'seguiraqui
    'manu, reimplemnta alignement
    'Alignement = lbList(0).Alignment
End Property

Public Property Let Alignment(val As AlignmentConstants)
    'seguiraqui
    'lbList(0).Alignment = val
    'ReAcomodar
End Property

Public Property Get Font() As StdFont
    'seguiraqui
    'Set Font = lbList(0).Font
End Property

Public Property Set Font(val As StdFont)
    'seguiraqui
'    Dim a As Long
'    For a = 0 To lbList.Count - 1
'        Set lbList(a).Font = val
'    Next a
'
'    Set lbTITULO.Font = Font
'
'    ReAcomodar
End Property

Private Sub ReAcomodar()  'cuando cambia la fuente o algo se debe reacomodar
    
End Sub

Public Sub PintarFondo(HDC_Pintar As Long)
    BitBlt UserControl.HDC, 0, 0, UserControl.Width / 15, UserControl.Height / 15, HDC_Pintar, 0, 0, SRCCOPY
    UserControl.Refresh
End Sub

Public Sub ImitarFondo(hdcPadre As Long, cX As Long, cY As Long)
    BitBlt UserControl.HDC, 0, 0, UserControl.Width / 15, UserControl.Height / 15, hdcPadre, cX, cY, SRCCOPY
    'BitBlt UserControl.hdc, 0, 0, UserControl.Width / 15, UserControl.Height / 15, hdcPadre, UserControl.CurrentX, UserControl.CurrentY, SRCCOPY
    UserControl.Refresh
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    i = ListaLista.DoClick_GetElementoIndex(CLng(x / 15), CLng(y / 15))
    UserControl.Refresh
End Sub

