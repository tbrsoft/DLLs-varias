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
Public Type PROPIs
    NameProp As String
    ValueProp As String
End Type

Private mPRs() As PROPIs

Public Sub AddProp(sName As String, sValue As String)
    Dim k As Long
    k = UBound(mPRs) + 1
    ReDim Preserve mPRs(k)
    mPRs(k).NameProp = sName
    mPRs(k).ValueProp = sValue
End Sub

Public Function GetProp_STR(sName As String) As String
    GetProp_STR = GetProp(sName).ValueProp
End Function

Public Function GetProp_LNG(sName As String) As Long
    GetProp_LNG = CLng(GetProp(sName).ValueProp)
End Function

Public Function GetProp_BOOL(sName As String) As Boolean
    GetProp_BOOL = CBool(CLng(GetProp(sName).ValueProp))
End Function

Public Function GetProp(sName As String) As PROPIs
    Dim k As Long
    For k = 1 To UBound(mPRs)
        If LCase(mPRs(k).NameProp) = LCase(sName) Then
            GetProp = mPRs(k)
            Exit Function
        End If
    Next k
    
    'error, se pide prop inexistente!
    
    'podemos buscar valores predeterminados para funciones comunes
    
    Dim PP As PROPIs
    PP = GetProDef(sName)
    terr.AppendLog "NoPR.,mnn", PP.NameProp + ":" + PP.ValueProp
    GetProp = PP
        
End Function

Private Function GetProDef(sNameProp As String) As PROPIs

    Dim PP As PROPIs
    PP.NameProp = sNameProp
    Select Case LCase(sNameProp)
        Case "fontname":            PP.ValueProp = "Verdana"
        Case "fontsize":            PP.ValueProp = "10"
        Case "fontbold":            PP.ValueProp = "0"
        Case "fontitalic":          PP.ValueProp = "0"
        Case "fontstrike":          PP.ValueProp = "0"
        Case "fintunderline":       PP.ValueProp = "0"
        Case "fontcolor":           PP.ValueProp = "455"
        Case "fontcolorsel":        PP.ValueProp = "2311"
        Case "fontcolorunsel":      PP.ValueProp = "6188"
        Case "fontcolorcontorno":   PP.ValueProp = "9899"
        Case "alphahabilitado":     PP.ValueProp = "1"
        Case "alphacolorlong":      PP.ValueProp = "5555"
        Case "alphacantidad":       PP.ValueProp = "128"
        
        Case Else
            terr.AppendSinHist "NOPROP_I:" + CStr(sNameProp)
            PP.ValueProp = "0" 'este sirve para boolean long y string
    End Select
    
    GetProDef = PP
    
End Function

Public Property Get X() As Long:   X = mX: End Property
Public Property Get Y() As Long:   Y = mY: End Property
Public Property Get W() As Long:   W = mW: End Property
Public Property Get H() As Long:   H = mH: End Property

Public Property Get X2() As Long:   X2 = mX2: End Property
Public Property Get Y2() As Long:   Y2 = mY2: End Property
Public Property Get W2() As Long:   W2 = mW2: End Property
Public Property Get H2() As Long:   H2 = mH2: End Property

Public Property Let X(newVal As Long):    mX = newVal:    End Property
Public Property Let Y(newVal As Long):    mY = newVal:    End Property
Public Property Let W(newVal As Long):    mW = newVal:    End Property
Public Property Let H(newVal As Long):    mH = newVal:    End Property

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
    ReDim mPRs(0)
End Sub

Public Sub DOIT(sComando As String)
    'le digo al manu que el objeto haga algun comando
    'SEGUIRAQUI
    'ejemplo "text=asa"
End Sub

Public Function ToString() As String
    Dim TMP As String
    TMP = "Name: " + SGOName + vbCrLf + _
          "  -rect: " + CStr(mX2) + " " + CStr(mY2) + " " + CStr(mW2) + " " + CStr(mH2) + vbCrLf + _
          "  -align: " + CStr(AlignementH) + " " + CStr(AlignementV) + vbCrLf + _
          "  -estirable: " + CStr(Estirable) + vbCrLf
          
    Dim M As Long
    For M = 1 To UBound(mPRs)
        TMP = TMP + "  -PROP: " + mPRs(M).NameProp + "=" + mPRs(M).ValueProp + vbCrLf
    Next M
    
    ToString = TMP
    
End Function
