VERSION 5.00
Begin VB.UserControl BsnVideoIN 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4275
   ScaleHeight     =   1575
   ScaleWidth      =   4275
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.ComboBox cmboSource 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   4275
   End
End
Attribute VB_Name = "BsnVideoIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Event Declarations:
'MappingInfo=UserControl,UserControl,-1,Click
  Event Click()
'MappingInfo=cmboSource,cmboSource,-1,DblClick
  Event DblClick()
'MappingInfo=cmboSource,cmboSource,-1,KeyDown
  Event KeyDown(KeyCode As Integer, Shift As Integer)
'MappingInfo=cmboSource,cmboSource,-1,KeyPress
  Event KeyPress(KeyAscii As Integer)
'MappingInfo=cmboSource,cmboSource,-1,KeyUp
  Event KeyUp(KeyCode As Integer, Shift As Integer)
'MappingInfo=UserControl,UserControl,-1,MouseDown
  Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MappingInfo=UserControl,UserControl,-1,MouseMove
  Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MappingInfo=UserControl,UserControl,-1,MouseUp
  Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub cmboSource_Click()
    RaiseEvent Click

End Sub

Private Sub UserControl_Initialize()
  Dim lpszName As String * 100
  Dim lpszVer As String * 100
  Dim X As Integer
  Dim lResult As Long
  Dim Caps As CAPDRIVERCAPS


    '// Get a list of all the installed drivers
    X = 0
  Do
        lResult = capGetDriverDescriptionA(X, lpszName, 100, lpszVer, 100)   '// Retrieves driver info
        If lResult Then
            cmboSource.AddItem lpszName
            X = X + 1
        End If
  Loop Until lResult = False

    '// Get the capabilities of the current capture driver
    lResult = capDriverGetCaps(lwndC, VarPtr(Caps), Len(Caps))

    '// Select the driver that is currently being used
    If lResult Then cmboSource.ListIndex = Caps.wDeviceIndex

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    cmboSource.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set cmboSource.Font = PropBag.ReadProperty("Font", Ambient.Font)
    cmboSource.Enabled = PropBag.ReadProperty("Enabled", True)
    cmboSource.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)

End Sub

Private Sub UserControl_Resize()

With cmboSource
   .Left = 0
   .Top = 0
   .Width = ScaleWidth + 1
End With

With UserControl
   .Width = cmboSource.Width + 1
   .Height = cmboSource.Height + 1
End With



End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", cmboSource.BackColor, &H80000005)
    Call PropBag.WriteProperty("Font", cmboSource.Font, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", cmboSource.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", cmboSource.ForeColor, &H80000008)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmboSource,cmboSource,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = cmboSource.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    cmboSource.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmboSource,cmboSource,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = cmboSource.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cmboSource.Font = New_Font
    PropertyChanged "Font"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmboSource,cmboSource,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = cmboSource.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    cmboSource.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmboSource,cmboSource,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = cmboSource.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    cmboSource.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmboSource,cmboSource,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    cmboSource.AddItem Item, Index
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=cmboSource,cmboSource,-1,ListIndex
'Public Property Get ListIndex() As Integer
'    ListIndex = cmboSource.ListIndex
'End Property
'
Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    cmboSource.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property
'
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=cmboSource,cmboSource,-1,ListIndex
'Public Property Get ListIndex() As Integer
'    ListIndex = cmboSource.ListIndex
'End Property
'
Private Sub cmboSource_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub cmboSource_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cmboSource_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cmboSource_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmboSource,cmboSource,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = cmboSource.ListIndex
End Property

