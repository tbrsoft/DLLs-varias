VERSION 5.00
Begin VB.UserControl tbrDISK 
   AccessKeys      =   "was"
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   MaskColor       =   &H00000000&
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   Begin VB.PictureBox P 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   5310
      ScaleHeight     =   2550
      ScaleWidth      =   2550
      TabIndex        =   0
      Top             =   480
      Width           =   2550
   End
   Begin VB.Shape SH 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   1275
      Index           =   0
      Left            =   1050
      Shape           =   2  'Oval
      Top             =   1290
      Width           =   1635
   End
End
Attribute VB_Name = "tbrDISK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Private mRadio As Long
Private mXX As Long, mYY As Long 'del centro en ObjToDraw!

Private Sub PINT()
    'ponerImgTransp "c:\cd3.jpg", True
    'Randomize
    'DibujarElipseDegradee Int(Rnd * 22222222), Int(Rnd * 22222222), 3, False
    ponerImgTransp "c:\cd4.bmp"
    Dim SS(5) As String
    SS(1) = "11111111111111"
    SS(2) = "22222222222222"
    SS(3) = "33333333333333"
    SS(4) = "44444444444444"
    SS(5) = "55555555555555"
    EscribirBorde "Verdana", True, 10, SS
    UserControl.Width = UserControl.Width + 10
End Sub

Private Sub UserControl_Click()
    PINT
End Sub

Private Sub UserControl_Initialize()

    UserControl.BackStyle = 1
    UserControl.AutoRedraw = True
    UserControl.ScaleMode = vbPixels
    
    'UserControl.Refresh
    
    P.Visible = False
    P.ScaleMode = vbPixels
    P.AutoSize = True
    P.AutoRedraw = True
    
    SH(0).BorderWidth = 2
    SH(0).Visible = False
    
End Sub

Public Sub ponerImgTransp(archIMG As String, Optional AutoSizeCtrl As Boolean = False)
    'el color transparente es el pixel 1,1 (x,y) de la imagen!
    DoEvents
    P.Picture = LoadPicture(archIMG)
    
    If AutoSizeCtrl Then
        UserControl.Width = P.Width
        UserControl.Height = P.Height
    End If
    TransparentBlt UserControl.hdc, 0, 0, _
        UserControl.ScaleWidth, UserControl.ScaleHeight, _
        P.hdc, 0, 0, _
        P.ScaleWidth, P.ScaleHeight, RGB(255, 255, 255) 'P.Point(1, 1)
        
    'UserControl.Refresh
End Sub

Public Sub EscribirBorde(lFontName As String, _
    lFontBold As Boolean, lFontSize As Long, Lista() As String)
    
    'escribir las letras!
    UserControl.Font.Name = lFontName
    UserControl.Font.Size = lFontSize
    UserControl.Font.Bold = lFontBold
    
    Dim J As Long
    'arranco un renglon abajo del tope
    UserControl.CurrentY = mYY - mRadio + UserControl.TextHeight("yY")
    For J = 1 To UBound(Lista)
        'ubicar las coordenadas segun corresponda!
        UserControl.CurrentY = UserControl.CurrentY + UserControl.TextHeight("yY")
        
        'creo que sale de la funcion x2+y2=Radio2?
        '==>x=raiz(radio2-y2)
        Dim PosX As Long, PosY As Long
        PosY = (mYY - UserControl.CurrentY)
        'valor del centro mXX-radio + valor como si fuera desde cero
        PosX = (mXX - mRadio) + (mRadio - Sqr(Abs((mRadio ^ 2) - (PosY ^ 2))))
        UserControl.CurrentX = PosX '+ J / 2
        UserControl.Print Lista(J);
    Next J
    
End Sub

Public Sub DibujarElipseDegradee(ColorDesde As Long, ColorHasta As Long, _
    AnchoLinea As Long, Optional HacerCLS As Boolean = True)
    
    'xx e yy son el centro!
        
    If HacerCLS Then UserControl.Cls
    
    'ver como son en RGB para ir pasando!
    Dim R1 As Long, R2 As Long
    Dim G1 As Long, G2 As Long
    Dim B1 As Long, B2 As Long
    
    B1 = ColorDesde \ 65536
    B2 = ColorHasta \ 65536
    
    G1 = (ColorDesde - (B1 * 65536)) \ 256
    G2 = (ColorHasta - (B2 * 65536)) \ 256
    
    R1 = (ColorDesde - (B1 * 65536) - (G1 * 256))
    R2 = (ColorHasta - (B2 * 65536) - (G2 * 256))
    
    Dim R3 As Long, G3 As Long, B3 As Long
    R3 = R1: G3 = G1: B3 = B1
    
    Dim VarR As Long, VarG As Long, VarB As Long
    VarR = (R2 - R1): VarG = (G2 - G1): VarB = (B2 - B1)
    
    'UserControl.DrawWidth = AnchoLinea
    SH(0).BorderWidth = AnchoLinea

    Dim J As Long
    Dim R4 As Long, G4 As Long, B4 As Long
    For J = 1 To mRadio 'Step AnchoLinea
        R4 = R3 + (VarR * (J / mRadio))
        G4 = G3 + (VarG * (J / mRadio))
        B4 = B3 + (VarB * (J / mRadio))
        On Local Error Resume Next
        Load SH(J)
        On Error GoTo 0
        SH(J).Width = J * 2: SH(J).Height = J * 2
        SH(J).Top = mYY - J: SH(J).Left = mXX - J
        SH(J).BorderColor = RGB(R4, G4, B4)
        SH(J).Visible = True
        'UserControl.Circle (mXX, mYY), J, RGB(R4, G4, B4)
        'UserControl.Print CStr(RGB(R4, G4, B4))
    Next J
    'UserControl.Refresh
End Sub

Public Sub Borrar()
    UserControl.Cls
End Sub

Private Sub UserControl_Resize()
    mRadio = UserControl.Height / 30 '2 'queda fuera del borde si lo pongo exacto!
    mXX = UserControl.Width / 30
    mYY = UserControl.Height / 30
End Sub

Private Sub UserControl_Terminate()
    P.Picture = LoadPicture
    Borrar
End Sub
