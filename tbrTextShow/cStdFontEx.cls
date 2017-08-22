VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStdFontEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ===========================================================================
'
' Filename:    cStdFontEx
' Author:      Steven Higgan
' Date:        9 November 2001
'
' Requires:    NOTHING
'
' Platform: Built on Windows 2000 using Visual Basic 6 Enterprise
'
' Tested: NOT TESTED ON ANY OTHER PLATFORMS
'
' Pourpose: used in cTextEx, mirrors functonalaty provided by the StdFont
' object. i only added a Colour Property (it makes sence to have do this)
' future revisions of this class will encorperate diffrent brush styles
' thus the reasoning for mirroring the vb font object
'
' Built on 9 Dec 2001 as part of research into my up coming Menu Object
' a full implementation of a Menu with no user controls, built compleatly
' from the ground up using the Windows API
'
' Other Peoples Code: None
'
' if you change this code, add more funcunalaty debug it on other windows
' platforms i would apreciate a copy of any changes made to the source code
'
' ===========================================================================

Private m_oStdFont As StdFont
Private m_lColour As Long
Private m_sName As String
Private m_lSize As Long
Private m_bItalic As Boolean
Private m_bBold As Boolean
Private m_bUnderline As Boolean
Private m_bStrikethrough As Boolean
Private m_lCharset As Long

Public Property Set StdFont(vDta As StdFont)
  Set m_oStdFont = vDta
  With m_oStdFont
    m_sName = .Name
    m_lSize = .Size
    m_bItalic = .Italic
    m_bBold = .Bold
    m_bUnderline = .UnderLine
    m_bStrikethrough = .Strikethrough
    m_lCharset = .Charset
  End With
End Property
Public Property Get StdFont() As StdFont
  If Not (m_oStdFont Is Nothing) Then
    Set StdFont = m_oStdFont
    Else
    'Raise Error - Not Set
  End If
End Property

Public Property Let Colour(vDta As Long)
  m_lColour = vDta
End Property
Public Property Let Name(vDta As String)
  m_sName = vDta
End Property
Public Property Let Size(vDta As Long)
  m_lSize = vDta
End Property
Public Property Let Italic(vDta As Boolean)
  m_bItalic = vDta
End Property
Public Property Let Bold(vDta As Boolean)
  m_bBold = vDta
End Property
Public Property Let UnderLine(vDta As Boolean)
  m_bUnderline = vDta
End Property
Public Property Let Strikethrough(vDta As Boolean)
  m_bStrikethrough = vDta
End Property
Public Property Let Charset(vDta As Long)
  m_lCharset = vDta
End Property

Public Property Get Colour() As Long
  If m_lColour <> 0 Then
  Colour = m_lColour
  Else
    'Raise Error - Not Set
    End If
End Property
Public Property Get Name() As String
  If m_lColour <> 0 Then
    Name = m_sName
    Else
    'Raise Error - Not Set
    End If
End Property
Public Property Get Size() As Long
  If m_lSize <> 0 Then
    Size = m_lSize
    Else
    'Raise Error - Not Set
    End If
End Property
Public Property Get Italic() As Boolean
  Italic = m_bItalic
End Property
Public Property Get Bold() As Boolean
  Bold = m_bBold
End Property
Public Property Get UnderLine() As Boolean
  UnderLine = m_bUnderline
End Property
Public Property Get Strikethrough() As Boolean
  Strikethrough = m_bStrikethrough
End Property
Public Property Get Charset() As Long
'TO DO  - _
        Change it so that it returns the Correct Charset in the enviroment that it _
        is operating in _
 _
        You could be a lazy bastard and simply loop thru the screen.font object _
        looking for the FontName and getting its .Charaset property but i know _
        that you all are real programmers and wouldnt stoop soo low - and in _
        all likelyhood cause a dramattic slow down on Slow Machines with Lots of _
        Fonts installed
        
  If m_lCharset <> 0 Then
    Charset = m_lCharset
    Else
    Charset = 0 'Better hope that its a Standard Windows character
    End If
End Property
