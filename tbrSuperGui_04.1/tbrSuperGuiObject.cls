VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrSuperGuiObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public SGOName As String
Public Estirable As Boolean

Private mAlignementH As Long '0 izquierda 1 centro 2 derecha
Private mAlignementV As Long '0 izquierda 1 centro 2 derecha

Private mX As Long
Private mY As Long
Private mW As Long
Private mH As Long
'traducido
Private mX2 As Long
Private mY2 As Long
Private mW2 As Long
Private mH2 As Long

'como no quiero reimplementar toooodas las propiedades
'le dejo generico que se maneje una lista de propiedades

Private mPRs() As clsPropis

Private Function AddProp(sName As String, sValue As String) As clsPropis
    'si la propiedad ya existe solo actualziarla!!!
    
    Dim k As Long
    k = UBound(mPRs) + 1
    ReDim Preserve mPRs(k)
    Set mPRs(k) = New clsPropis
    mPRs(k).NameProp = sName
    mPRs(k).ValueProp = sValue
    
    Set AddProp = mPRs(k)
End Function

Public Function GetPropCantidad() As Long
    GetPropCantidad = UBound(mPRs)
End Function

Public Function GetPropByID(Ix As Long) As clsPropis
    If Ix = 0 Or Ix > UBound(mPRs) Then
        Set GetPropByID = mPRs(0)
        terr.AppendLog "NoProp:" + CStr(Ix)
    Else
        Set GetPropByID = mPRs(Ix)
    End If
End Function

Public Function GetProp_STR(sName As String) As String
    GetProp_STR = GetProp(sName).ValueProp
End Function

Public Function GetProp_LNG(sName As String) As Long
    GetProp_LNG = CLng(GetProp(sName).ValueProp)
End Function

Public Function GetProp_BOOL(sName As String) As Boolean
    GetProp_BOOL = CBool(CLng(GetProp(sName).ValueProp))
End Function

'esta se usa mas con la intencion de obtener un valor
Public Function GetProp(sName As String) As clsPropis
    
    Dim k As Long
    For k = 1 To UBound(mPRs)
        If LCase(mPRs(k).NameProp) = LCase(sName) Then
            Set GetProp = mPRs(k)
            Exit Function
        End If
    Next k
    
    'podemos buscar valores predeterminados para funciones comunes
    Set GetProp = GetProDef(sName)
        
End Function

Public Sub SetProp(sName As String, sValue As String)
    'cambiar el valor de una variable y crearla si no existe
    '????????las tres lineas que siguen no funcionaban ?????????????
    'Dim P As PROPIs
    'P = GetPropOrCreate(sName, sValue)
    'P.ValueProp = sValue
    
    Dim k As Long
    For k = 1 To UBound(mPRs)
        If LCase(mPRs(k).NameProp) = LCase(sName) Then
            mPRs(k).ValueProp = sValue
            Exit Sub
        End If
    Next k
    
    AddProp sName, sValue
    
End Sub

Public Function GetPropOrCreate(sName As String, Optional defValue As String = "0") As clsPropis
    Dim k As Long
    For k = 1 To UBound(mPRs)
        If LCase(mPRs(k).NameProp) = LCase(sName) Then
            Set GetPropOrCreate = mPRs(k)
            Exit Function
        End If
    Next k
    
    Dim PP As clsPropis
    PP = AddProp(sName, defValue)
    
    Set GetPropOrCreate = PP
        
End Function


Private Function GetProDef(sNameProp As String) As clsPropis

    Dim PP As New clsPropis
    PP.NameProp = sNameProp
    Select Case LCase(sNameProp)
        Case "fontname":            PP.ValueProp = "Verdana"
        Case "fontsize":            PP.ValueProp = "10"
        Case "fontbold":            PP.ValueProp = "0"
        Case "fontitalic":          PP.ValueProp = "0"
        Case "fontstrike":          PP.ValueProp = "0"
        Case "fontunderline":       PP.ValueProp = "0"
        Case "fontcolor":           PP.ValueProp = "22455"
        Case "fontcolorsel":        PP.ValueProp = "2311"
        Case "fontcolorunsel":      PP.ValueProp = "6188"
        Case "fntcolorcontorno":    PP.ValueProp = "9899"
        Case "alphahabilitado":     PP.ValueProp = "1"
        Case "alpahcolorlong":      PP.ValueProp = "5555" 'esta mal escrita por que me equivoque y en todos los demas lugares esta asi
        Case "alphacantidad":       PP.ValueProp = "128"
        Case "pixelsalteo":         PP.ValueProp = "3" 'seguiraqui AGREGAR a frmEDITMANU
        Case "tiemposentretextos":  PP.ValueProp = "40" 'seguiraqui AGREGAR a frmEDITMANU
        Case "colorsel":            PP.ValueProp = "5487" 'barra que selecciona cate temalist en un tema manager
        Case "colodiscorsel":       PP.ValueProp = "15487" 'barra que selecciona cate temalist en un tema manager
        
        Case Else
            terr.AppendSinHist "NOPROP_I:" + CStr(sNameProp)
            PP.ValueProp = "0" 'este sirve para boolean long y string
    End Select
    
    Set GetProDef = PP
    
End Function

Public Property Get X() As Long:   X = mX: End Property
Public Property Get Y() As Long:   Y = mY: End Property
Public Property Get W() As Long:   W = mW: End Property
Public Property Get H() As Long:   H = mH: End Property

Public Property Get X2() As Long:   X2 = mX2: End Property
Public Property Get Y2() As Long:   Y2 = mY2: End Property
Public Property Get W2() As Long:   W2 = mW2: End Property
Public Property Get H2() As Long:   H2 = mH2: End Property

'mientras no se traduzac x2 dee ser x
Public Property Let X(newVal As Long)
    mX = newVal
    mX2 = mX
End Property

Public Property Let Y(newVal As Long):    mY = newVal:    mY2 = mY: End Property
Public Property Let W(newVal As Long):    mW = newVal:    mW2 = mW: End Property
Public Property Let H(newVal As Long):    mH = newVal:    mH2 = mH: End Property

Public Property Let X2(newVal As Long):    mX2 = newVal:    End Property
Public Property Let Y2(newVal As Long):    mY2 = newVal:    End Property
Public Property Let W2(newVal As Long):    mW2 = newVal:    End Property
Public Property Let H2(newVal As Long):    mH2 = newVal:    End Property

Public Property Get AlignementH() As Long
    AlignementH = mAlignementH
End Property

Public Property Let AlignementH(newV As Long)
    mAlignementH = newV
End Property

Public Property Get AlignementV() As Long
    AlignementV = mAlignementV
End Property

Public Property Let AlignementV(newV As Long)
    mAlignementV = newV
End Property


Private Sub Class_Initialize()
    'predeterminados
    mAlignementH = 1
    mAlignementV = 1
    ClearPropis
End Sub

Public Sub ClearPropis()
    ReDim mPRs(0)
End Sub

Public Sub DOIT(sComando As String)
    'le digo al manu que el objeto haga algun comando
    'SEGUIRAQUI
    'ejemplo "text=asa"
End Sub

Public Function toString() As String
    Dim TMP As String
    TMP = "Name: " + SGOName + vbCrLf + _
          "  -rect: " + CStr(mX2) + " " + CStr(mY2) + " " + CStr(mW2) + " " + CStr(mH2) + vbCrLf + _
          "  -align: " + CStr(AlignementH) + " " + CStr(AlignementV) + vbCrLf + _
          "  -estirable: " + CStr(Estirable) + vbCrLf
          
    Dim M As Long
    For M = 1 To UBound(mPRs)
        TMP = TMP + "  -PROP: " + mPRs(M).NameProp + "=" + mPRs(M).ValueProp + vbCrLf
    Next M
    
    toString = TMP
    
End Function

'transportar coordenadas y tamaño segun pedido
Public Sub Traslate(Optional WidthScreenActual As Long = -1, Optional HeightScreenActual As Long = -1, _
    Optional WiBasePXUsado As Long = -1, Optional HeBasePXUsado As Long = -1)
    
    On Local Error Resume Next
    
    '*******************************************
    '*******************************************
    'referencias de a que pantalla quiero llegar
    Dim wi As Long, he As Long
    If WidthScreenActual = -1 Then
        wi = Screen.Width / 15
    Else
        wi = WidthScreenActual
    End If
    
    If HeightScreenActual = -1 Then
        he = Screen.Height / 15
    Else
        he = HeightScreenActual
    End If
    
    '*******************************************
    '*******************************************
    'valores usados por todo el skin de manera predeterminada
    Dim WiBasePX As Long, HeBasePX As Long

    If WiBasePXUsado = -1 Then
        WiBasePX = HechoParaPixlesAncho
    Else
        WiBasePX = WiBasePXUsado
    End If
    
    If HeBasePXUsado = -1 Then
        HeBasePX = HechoParaPixlesAlto
    Else
        HeBasePX = HeBasePXUsado
    End If
    
    '*******************************************
    '*******************************************
    Dim coefW As Double
    Dim coefH As Double
    
    coefW = wi / WiBasePX
    coefH = he / HeBasePX
    
    'aplicar a "2"

    If OB.Estirable Then 'joiaaaaaaaaaaaaa
        mX2 = CLng(CDbl(mX) * coefW)
        mY2 = CLng(CDbl(mY) * coefH)
        mW2 = CLng(CDbl(mW) * coefW)
        mH2 = CLng(CDbl(mH) * coefH)
    Else
        'si no es estirable debe correr su x e y para acomodarse y seguir mas o menos centrado en su zona
        mX2 = CLng(CDbl(mX) * coefW)
        mY2 = CLng(CDbl(mY) * coefH)
        
        mW2 = mW
        mH2 = mH
        
        'ver que alineacion tiene!
        Dim ExtraX As Long
        Dim ExtraY As Long
        
        'alineacion centrada predeterminada
        'movieminto extra (para que quede centrado en la zona que ocuparia el objeto si fuera estirable. S O Y   U N   G E N I O
        If mAlignementH = 1 Then ExtraX = ((CLng(CDbl(mW) * coefW)) - mW) / 2
        If mAlignementV = 1 Then ExtraY = ((CLng(CDbl(mH) * coefH)) - mH) / 2
        'izquierda en H y arriba en V
        'If ob.AlignementH = 0 Then 'izquierda (nada, ya esta ahi!!)
        'If ob.AlignementV = 0 Then 'arriba (nada, ya esta ahi!!)
        'derecha en H y abajo en Y
        If mAlignementH = 2 Then ExtraX = ((CLng(CDbl(mW) * coefW)) - mW)
        If mAlignementV = 2 Then ExtraY = ((CLng(CDbl(mH) * coefH)) - mH)
        
        mX2 = mX2 + ExtraX
        mY2 = mY2 + ExtraY
        
    End If
    
    '... y ya se puede usar x2, y2, etc
End Sub
