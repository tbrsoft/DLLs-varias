VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTextEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
  Private Const BITSPIXEL = 12
  Private Const LOGPIXELSX = 88
  Private Const LOGPIXELSY = 90

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
  Private Const NEWTRANSPARENT = 3

Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
  Private Const CLR_INVALID = &HFFFF

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
  Private Const DT_BOTTOM = &H8
  Private Const DT_CENTER = &H1
  Private Const DT_LEFT = &H0
  Private Const DT_CALCRECT = &H400
  Private Const DT_WORDBREAK = &H10
  Private Const DT_VCENTER = &H4
  Private Const DT_TOP = &H0
  Private Const DT_TABSTOP = &H80
  Private Const DT_SINGLELINE = &H20
  Private Const DT_RIGHT = &H2
  Private Const DT_NOCLIP = &H100
  Private Const DT_INTERNAL = &H1000
  Private Const DT_EXTERNALLEADING = &H200
  Private Const DT_EXPANDTABS = &H40
  Private Const DT_CHARSTREAM = 4
  Private Const DT_NOPREFIX = &H800

Private Const LF_FACESIZE = 32
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type


'(((((((((((((((((((((((((((())))))))))))))))))))))))))))))
Public Enum eFormatFlags
  Bottom = DT_BOTTOM
  center = DT_CENTER
  Left = DT_LEFT
  CALCRECT = DT_CALCRECT
  WORDBREAK = DT_WORDBREAK
  vCENTER = DT_VCENTER
  Top = DT_TOP
  TabStop = DT_TABSTOP
  SINGLELINE = DT_SINGLELINE
  Right = DT_RIGHT
  NOCLIP = DT_NOCLIP
  INTERNAL = DT_INTERNAL
  EXTERNALLEADING = DT_EXTERNALLEADING
  EXPANDTABS = DT_EXPANDTABS
  CHARSTREAM = DT_CHARSTREAM
  NOPREFIX = DT_NOPREFIX
End Enum

Private m_tTextRect As RECT
Private m_oStdFontEx As cStdFontEx
Private m_lhDc As Long
Private m_sText As String
Private m_lAllignmentFlags As eFormatFlags

Public Property Set StdFontEx(vDta As cStdFontEx)
  Set m_oStdFontEx = vDta
End Property
Public Property Get StdFontEx() As cStdFontEx
  If Not (m_oStdFontEx Is Nothing) Then
    Set StdFontEx = m_oStdFontEx
    End If
End Property

Public Property Let RectBottom(vDta As Long)
  m_tTextRect.Bottom = vDta
End Property
Public Property Let RectTop(vDta As Long)
  m_tTextRect.Top = vDta
End Property
Public Property Let RectLeft(vDta As Long)
  m_tTextRect.Left = vDta
End Property
Public Property Let RectRight(vDta As Long)
  m_tTextRect.Right = vDta
End Property

Public Property Let hdc(vDta As Long)
  m_lhDc = vDta
End Property
Public Property Let Text(vDta As String)
  m_sText = vDta
End Property
Public Property Let AllignmentFlags(ByVal vDta As eFormatFlags)
  m_lAllignmentFlags = vDta
End Property

Public Sub Draw(Optional oStdFontEx As cStdFontEx = Nothing, _
                Optional RectBottom As Long = -1, _
                Optional RectTop As Long = -1, _
                Optional RectLeft As Long = -1, _
                Optional RectRight As Long = -1, _
                Optional hdc As Long = -1, _
                Optional Text As String = "", _
                Optional ByRef AllignmentFlags As eFormatFlags = -1)
                
  Dim LF As LOGFONT
  Dim sFont As String
  Dim iChar As Integer
  Dim m_lFont As Long
  Dim crOldType As Long
  
  Dim ObjetoTmp As Long
  
  '//Check to see whether the user passed any arguements in the .Draw procedure _
     if no arguements were passed then use the Class.* propertys
  If Not (oStdFontEx Is Nothing) Then
    Set m_oStdFontEx = oStdFontEx
    End If
    
  If Not (RectBottom = -1) Then
    m_tTextRect.Bottom = RectBottom
    End If
  If Not (RectTop = -1) Then
    m_tTextRect.Top = RectTop
    End If
  If Not (RectLeft = -1) Then
    m_tTextRect.Left = RectLeft
    End If
  If Not (RectRight = -1) Then
    m_tTextRect.Bottom = RectRight
    End If
    
  If Not (hdc = -1) Then
    m_lhDc = hdc
    End If
  If Not (Text = "") Then
    m_sText = Text
    End If
  If Not (AllignmentFlags = -1) Then
    m_lAllignmentFlags = AllignmentFlags
    End If
  
  '//Convert StdFontEx to the LogFont Structure so that it can be applied to the hDc
  With LF
    sFont = m_oStdFontEx.Name
    ' There is a quicker way involving StrConv and CopyMemory, but
    ' this is simpler!:
    For iChar = 1 To Len(sFont)
      .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
      Next iChar
    ' Based on the Win32SDK documentation:
    .lfHeight = -MulDiv((m_oStdFontEx.Size), (GetDeviceCaps(m_lhDc, LOGPIXELSY)), 72)
    .lfItalic = m_oStdFontEx.Italic
    If (m_oStdFontEx.Bold) Then
      .lfWeight = FW_BOLD
      Else
        .lfWeight = FW_NORMAL
      End If
    .lfUnderline = m_oStdFontEx.UnderLine
    .lfStrikeOut = m_oStdFontEx.Strikethrough
    .lfCharSet = m_oStdFontEx.Charset
  End With
  
  '//Apply the Font Formatting to the hDc
  m_lFont = CreateFontIndirect(LF)
  
  ObjetoTmp = SelectObject(m_lhDc, m_lFont)
  ret = DeleteObject(ObjetoTmp)
  ret = DeleteObject(m_lFont)

  '//Apply the Text Colour Formatting to the hDc
  SetBkMode m_lhDc, NEWTRANSPARENT
  SetTextColor m_lhDc, m_oStdFontEx.Colour
  
  '//OutPut the Text
  DrawText m_lhDc, m_sText, Len(m_sText), m_tTextRect, m_lAllignmentFlags
  
End Sub
