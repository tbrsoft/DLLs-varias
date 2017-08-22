VERSION 5.00
Begin VB.UserControl MsgScroller 
   CanGetFocus     =   0   'False
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   ForeColor       =   &H8000000E&
   MouseIcon       =   "MsgScroller.ctx":0000
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   34
   ToolboxBitmap   =   "MsgScroller.ctx":0152
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   180
      Left            =   45
      Top             =   45
   End
End
Attribute VB_Name = "MsgScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Message Scroller Control
'
' This control is a message displaying control which
' allows multiple messages to be displayed in a scrolling
' presentation.  Messages can be Added or Deleted from
' the control while messages are being displayed.
'
' Methods:
'   1) AddItem - Allows the addition of new messages into the message queue.
'   2) RemoveItem - Removal of a message from the message queue.
'
' Properties:
'   1) Border - Set whether or not the control has a border. True for sunken
'      border, False for no border.
'   2) Backcolor - Set the backcolor of the control.
'   3) ForeColor - Set the color of the message text.
'   4) Font - Set the type and size of the font for the messages.
'   5) ScrollInterval - Time in milliseconds between movements of the message
'      text in the scroll window.
'   6) ScrollSpeed - Length in pixels of each movement within the scroll window.
'   7) Scroll - Starts/Stops the scrolling of the messages.
'
Option Explicit

Public Enum SPBorderStyle
    [None] = 0
    [Fixed Single] = 1
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Default Property Values:
Const m_def_Border = [Fixed Single]
Const m_def_ScrollSpeed = 3             ' Pixels to move each tick
Const m_def_ScrollInterval = 120        ' Milliseconds
Const m_def_Seperator As String = "·"
Const m_def_Enabled = 0
Const m_def_AutoScroll As Boolean = False
Const m_def_Pan As Boolean = True

' Property Variables:
Private m_ScrollSpeed As Integer
Private m_Enabled As Boolean
Private m_Seperator As String
Private m_AutoScroll As Boolean
Private m_Pan As Boolean

' Interval Variables:
Private Message As New Collection
Private mnCurrentIndex As Integer
Private mbForceScroll As Boolean
Private mbAllowScroll As Boolean
Private mbStopWhenReady As Boolean
Private mnGroupCount As Integer
Private msSpacer As String
Private mnTotalWidth As Long
Private mnSpacerWidth As Long
Private mnPosition As Long

' Panning vars
Private Type PAN_DATA
    Moving As Boolean
    InitX As Single
    InitPosition As Long
    TimerActive As Boolean
End Type
Private mPan As PAN_DATA

' Mouse sensitive areas
Private Type HOT_ITEM_DATA
    Index As Long
    LeftX As Long
    RightX As Long
End Type
Private Type HOT_DATA
    Item() As HOT_ITEM_DATA
    Count As Long
    TopY As Long
    BottomY As Long
    Spot As Long
End Type
Private mHot As HOT_DATA

Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_SINGLELINE = &H20
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800

Private Const OPAQUE As Long = 2

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Event HotSpotClick(ByVal Index As Long)
Event HotSpotMove(ByVal Index As Long)
Event Click()
Event DblClick()
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Attribute MouseUp.VB_UserMemId = -607
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Attribute MouseMove.VB_UserMemId = -606

Private Sub UserControl_InitProperties()
    'Initialize Properties for User Control

    m_Enabled = m_def_Enabled

    mbForceScroll = False
    mbAllowScroll = False
    mbStopWhenReady = False
    mnGroupCount = 0
    mnTotalWidth = 0&
    mnSpacerWidth = 0&
    mnPosition = 0&

    Erase mHot.Item
    mHot.Count = 0&

    m_AutoScroll = m_def_AutoScroll
    m_ScrollSpeed = m_def_ScrollSpeed
    m_Seperator = m_def_Seperator
    m_Pan = m_def_Pan

    UserControl.BorderStyle = m_def_Border
    UserControl.BackColor = &H8000000F
    UserControl.ForeColor = &H80000012
    tmrScroll.Interval = m_def_ScrollInterval

    SetControlState
End Sub

Private Sub UserControl_Paint()
    If Not tmrScroll.Enabled Then PaintControl
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("Border", m_def_Border)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    tmrScroll.Interval = PropBag.ReadProperty("ScrollInterval", m_def_ScrollInterval)

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)

    m_ScrollSpeed = PropBag.ReadProperty("ScrollSpeed", m_def_ScrollSpeed)
    m_Seperator = PropBag.ReadProperty("Seperator", m_def_Seperator)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_AutoScroll = PropBag.ReadProperty("AutoScroll", m_def_AutoScroll)
    m_Pan = PropBag.ReadProperty("Pan", m_def_Pan)

'    SetControlState
'    If Not Ambient.UserMode Then PaintControl
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Border", UserControl.BorderStyle, m_def_Border)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("ScrollInterval", tmrScroll.Interval, m_def_ScrollInterval)

    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)

    Call PropBag.WriteProperty("ScrollSpeed", m_ScrollSpeed, m_def_ScrollSpeed)
    Call PropBag.WriteProperty("Seperator", m_Seperator, m_def_Seperator)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("AutoScroll", m_AutoScroll, m_def_AutoScroll)
    Call PropBag.WriteProperty("Pan", m_Pan, m_def_Pan)
End Sub

Private Sub UserControl_Resize()
    SetControlState
    If Not Ambient.UserMode Then PaintControl
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        If m_Pan And mbAllowScroll Then
            With mPan
                .InitX = x
                .Moving = True
                .TimerActive = tmrScroll.Enabled
            End With
            tmrScroll.Enabled = False
            mPan.InitPosition = mnPosition
            MousePointer = vbCustom
        End If
    End If

    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mPan.Moving Then
        mnPosition = mPan.InitPosition - (x - mPan.InitX)
        PaintControl
    End If

    mHot.Spot = GetHotSpot(x, Y)
    RaiseEvent HotSpotMove(mHot.Spot)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mPan.Moving Then
        With mPan
            .Moving = False
            tmrScroll.Enabled = .TimerActive
        End With
        MousePointer = vbDefault
    End If

    mHot.Spot = GetHotSpot(x, Y)

    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_Click()
    If mHot.Spot >= 0 Then RaiseEvent HotSpotClick(mHot.Spot)
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

' *******************************************************************************************
' Designtime Properties

Public Property Get BorderStyle() As SPBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -502
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_Border As SPBorderStyle)
    UserControl.BorderStyle = New_Border
    PropertyChanged "Border"
    If Not Ambient.UserMode Then PaintControl
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    If Not Ambient.UserMode Then PaintControl
End Property

Public Property Get ScrollInterval() As Integer
Attribute ScrollInterval.VB_Description = "Returns/sets the number of milliseconds when to scroll the message."
Attribute ScrollInterval.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ScrollInterval = tmrScroll.Interval
End Property
Public Property Let ScrollInterval(ByVal New_ScrollInterval As Integer)
    tmrScroll.Interval = New_ScrollInterval
    PropertyChanged "ScrollInterval"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    SetControlState
    If Not Ambient.UserMode Then PaintControl
End Property

Public Property Get ScrollSpeed() As Integer
Attribute ScrollSpeed.VB_Description = "Returns/sets the number of pixels to scroll each time scroll event is called."
Attribute ScrollSpeed.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ScrollSpeed = m_ScrollSpeed
End Property
Public Property Let ScrollSpeed(ByVal New_ScrollSpeed As Integer)
    If New_ScrollSpeed < 1 Then Exit Property
    m_ScrollSpeed = New_ScrollSpeed
    PropertyChanged "ScrollSpeed"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Seperator() As String
Attribute Seperator.VB_Description = "Sets/returns characters that will seperate messages"
Attribute Seperator.VB_ProcData.VB_Invoke_Property = ";Text"
    Seperator = m_Seperator
End Property
Public Property Let Seperator(ByVal vNewValue As String)
    m_Seperator = vNewValue
    PropertyChanged "Seperator"
    SetControlState
    If Not Ambient.UserMode Then PaintControl
End Property

Public Property Get AutoScroll() As Boolean
Attribute AutoScroll.VB_Description = "Automatic starts the scrolling when total message length does not fit current view width."
Attribute AutoScroll.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoScroll = m_AutoScroll
End Property
Public Property Let AutoScroll(ByVal vNewValue As Boolean)
    m_AutoScroll = vNewValue
    PropertyChanged "AutoScroll"
    SetControlState
End Property

Public Property Get Scroll() As Boolean
Attribute Scroll.VB_Description = "Enabled/disabled message scrolling."
Attribute Scroll.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Scroll = mbForceScroll
End Property
Public Property Let Scroll(ByVal vNewValue As Boolean)
    mbForceScroll = vNewValue
    If Not Ambient.UserMode Then
        If Not mbForceScroll Then tmrScroll.Enabled = False
    End If
    SetControlState
End Property

Public Property Get AllowPan() As Boolean
Attribute AllowPan.VB_Description = "Returns/sets the state which allows users to pan the text using the mouse ."
Attribute AllowPan.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowPan = m_Pan
End Property
Public Property Let AllowPan(ByVal NewValue As Boolean)
    m_Pan = NewValue
    PropertyChanged "Pan"
End Property

' *******************************************************************************************
' Internal methods

Private Sub PaintControl()
    Dim nMaxWidth As Long, i As Long, nCount As Long, _
        hWinDC As Long, hObject As Long, nTextColor As Long, _
        nLeft As Long, nBegin As Long, nHeight As Long, _
        nTextTop As Long, nTextBottom As Long
    Dim tRect As RECT, tFill(0 To 1) As RECT
    Dim sMessage As String
    Dim bCheckStop As Boolean, bStop As Boolean

    On Local Error Resume Next

    ' Bounds-check required position
    If mbAllowScroll Then
        nMaxWidth = mnTotalWidth + mnSpacerWidth
    Else
        nMaxWidth = mnTotalWidth
    End If

    If mnPosition < 0 Then
        Do While mnPosition < 0                 ' Needed for panning
            mnPosition = mnPosition + nMaxWidth
        Loop
    Else
        Do While mnPosition > nMaxWidth
            mnPosition = mnPosition - nMaxWidth
        Loop
    End If

    bCheckStop = (mbStopWhenReady Or (Not mbAllowScroll))
    bStop = False

    ReDim mHot.Item(1 To 1024) As HOT_ITEM_DATA
    mHot.Count = 0&


    Call GetClientRect(hWnd, tRect)
    nMaxWidth = tRect.Right                     ' Width of view area
    mHot.TopY = 0&
    mHot.BottomY = tRect.Bottom

    If Not mbStopWhenReady Then mnGroupCount = 0

    hWinDC = hdc                                ' Get the Device Context handle value

    nHeight = (tRect.Bottom - TextHeight("X")) \ 2
    If nHeight > 0 Then
        With tFill(0)
            .Top = 0
            .Left = 0
            .Right = tRect.Right
            .Bottom = nHeight
        End With
        With tFill(1)
            .Top = nHeight + TextHeight("X")
            .Left = 0
            .Right = tRect.Right
            .Bottom = tRect.Bottom
        End With

        mHot.TopY = tFill(0).Bottom + 1
        mHot.BottomY = tFill(1).Top - 1

        hObject = CreateSolidBrush(GetColor(UserControl.BackColor))
        Call FillRect(hWinDC, tFill(0), hObject)
        Call FillRect(hWinDC, tFill(1), hObject)
        Call DeleteObject(hObject)
    End If

    Call SetBkMode(hWinDC, OPAQUE)              ' Make sure text is painted opaque (= not transparent)

    If Ambient.UserMode Then                    ' Runtime mode
        nCount = Message.Count                  ' Faster access when you plop it in a memvar
        If nCount = 0 Then                      ' No Messages available
            Cls
            Exit Sub
        End If

        nBegin = 1
        nLeft = mnPosition

        If nCount > 1 Then
            ' Using "mnPosition", find out where to start
            For i = 1 To nCount
                If mnPosition < Message.Item(i).Right Then
                    nBegin = i
                    nLeft = mnPosition - Message.Item(i).Left
                    Exit For
                End If
            Next
        End If

        i = nBegin
        tRect.Left = -(nLeft)

        Do While tRect.Left < nMaxWidth
            If Message.Item(i).Colour = -1 Then
                nTextColor = GetColor(UserControl.ForeColor)
            Else
                nTextColor = GetColor(Message.Item(i).Colour)
            End If

            ' Draw the text
            tRect.Right = tRect.Left + Message.Item(i).Width
            sMessage = Message.Item(i).Text
            '
            If mHot.Count < 1024 Then
                mHot.Count = mHot.Count + 1
                With mHot.Item(mHot.Count)
                    .Index = i
                    .LeftX = tRect.Left
                    .RightX = tRect.Right
                End With
            End If
            '
            hObject = SetTextColor(hWinDC, nTextColor)
            Call DrawText(hWinDC, sMessage, Len(sMessage), tRect, DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX)
            Call SetTextColor(hWinDC, hObject)
            '
            tRect.Left = tRect.Right

            If Not mPan.Moving Then

                If i = nCount Then
                    If Not mbAllowScroll Then
                        tRect.Right = tFill(0).Right
                        hObject = CreateSolidBrush(GetColor(UserControl.BackColor)) ' GetColor(UserControl.BackColor))
                        Call FillRect(hWinDC, tRect, hObject)
                        Call DeleteObject(hObject)
                        Exit Do
                    End If
                End If
            End If

            ' Draw the spacer
            tRect.Right = tRect.Left + mnSpacerWidth
            Call DrawText(hWinDC, msSpacer, Len(msSpacer), tRect, DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX)
            tRect.Left = tRect.Right

            i = i + 1
            If i > nCount Then i = 1
        Loop

    Else
        tRect.Left = -(mnPosition)
        sMessage = "Message Scroller" & msSpacer

        Do While tRect.Left < nMaxWidth
            tRect.Right = tRect.Left + mnTotalWidth + mnSpacerWidth
            Call DrawText(hWinDC, sMessage, Len(sMessage), tRect, DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX)
            tRect.Left = tRect.Right
        Loop
    End If

    If mHot.Count = 0& Then
        Erase mHot.Item
    Else
        ReDim Preserve mHot.Item(1 To mHot.Count) As HOT_ITEM_DATA
    End If
End Sub

Private Sub tmrScroll_Timer()
    mnPosition = mnPosition + m_ScrollSpeed
    PaintControl
End Sub

Private Function GetColor(ByVal nColor As Long) As Long
    Const SYSCOLOR_BIT As Long = &H80000000
    If (nColor And SYSCOLOR_BIT) = SYSCOLOR_BIT Then
        nColor = nColor And (Not SYSCOLOR_BIT)
        GetColor = GetSysColor(nColor)
    Else
        GetColor = nColor
    End If
End Function

Private Sub GetTotalWidth()
    Dim i As Long, nCount As Long, nWidth As Long

    If m_Seperator = "" Then                ' No spacer, force a space
        msSpacer = " "
    ElseIf Trim$(m_Seperator) = "" Then     ' Does it contain spaces only?
        msSpacer = m_Seperator
    Else                                    ' User-defined spacer
        msSpacer = " " & m_Seperator & " "
    End If

    mnSpacerWidth = TextWidth(msSpacer)
    mnTotalWidth = 0&

    If Ambient.UserMode Then                    ' Runtime mode
        nCount = Message.Count                  ' Faster access when you plop it in a memvar
        If nCount > 0 Then                      ' Messages available
            For i = 1 To nCount                 ' Calculate width as positions of each text block
                nWidth = TextWidth(Message.Item(i).Text)
                Message.Item(i).Width = nWidth
                Message.Item(i).Left = mnTotalWidth
                mnTotalWidth = mnTotalWidth + (nWidth + IIf(i < nCount, mnSpacerWidth, 0))
                Message.Item(i).Right = mnTotalWidth
            Next
        End If
    Else
        mnTotalWidth = TextWidth("Message Scroller")
    End If
End Sub

Private Sub SetControlState()
    Dim nCount As Long

    GetTotalWidth

    If mPan.Moving Then Exit Sub

    If Ambient.UserMode Then                ' Runtime mode
        nCount = Message.Count
    Else
        nCount = 1
    End If

    If nCount = 0 Then                      ' No more messages. Turn off
        If tmrScroll.Enabled Then           ' Ticker active
            tmrScroll.Enabled = False
        End If
        mnPosition = 0&
        Cls
    Else                                    ' Has a message
        If mbForceScroll Then               ' Has permission to scroll (forced)
            mbAllowScroll = True
        ElseIf m_AutoScroll Then            ' Scrolling depends on the message
            mbAllowScroll = (mnTotalWidth > ScaleWidth)
        Else                                ' Scrolling depends on the timer
            mbAllowScroll = False
        End If

        If mbAllowScroll Then
            mbStopWhenReady = False
            If Not tmrScroll.Enabled Then tmrScroll.Enabled = True  ' Ticker not active, activate it.
        Else                                                        ' Ticker not required
            tmrScroll.Enabled = False
            mnPosition = 0&
            PaintControl
        End If
    End If
End Sub

Private Function GetHotSpot(x As Single, Y As Single) As Long
    GetHotSpot = -1&

    If mHot.Count < 1 Then Exit Function
    If Y < mHot.TopY Then Exit Function
    If Y > mHot.BottomY Then Exit Function

    Dim i As Long

    For i = 1 To mHot.Count
        With mHot.Item(i)
            If x >= .LeftX Then
                If x <= .RightX Then
                    GetHotSpot = mHot.Item(i).Index
                    Exit For
                End If
            End If
        End With
    Next
End Function

Private Function ValidIndexType(Index As Variant) As Boolean
    Select Case VarType(Index)
    Case vbInteger, vbLong, vbString
        ValidIndexType = True
    Case Else
        ValidIndexType = False
    End Select
End Function

' *******************************************************************************************
' Public Properties/Methods

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
    SetControlState
    PaintControl
End Sub

' ToggleScroll method. This starts or stops the scrolling of the messages in the message queue.
Public Function ToggleScroll()
    If Not mPan.Moving And mbAllowScroll Then tmrScroll.Enabled = Not tmrScroll.Enabled
End Function

' Add message to Message collection.
Public Sub AddItem(Text As String, Optional Key As String = "", Optional Index As Long = 0)
Attribute AddItem.VB_Description = "Adds an item to the message queue."

    If Trim$(Text) = "" Then Exit Sub   ' Must have a string

    On Local Error GoTo Add_Error

    Dim tMessage As New cMessageData

    tMessage.Text = Text

    If Trim$(Key) = "" Then
        If Index > 0 Then
            Message.Add tMessage, , Index
        Else
            Message.Add tMessage
        End If
    ElseIf Index > 0 Then
        Message.Add tMessage, Key, Index
    Else
        Message.Add tMessage, Key
    End If

    SetControlState

Add_Error:
End Sub

' Find the text - return the index (0 = not found)
Public Function FindItem(ByVal Text As String) As Long
Attribute FindItem.VB_Description = "Finds an item by item text string. Returns the index. Zero when not found."
    Dim i As Integer

    On Local Error GoTo Find_Error

    If Message.Count = 0 Then GoTo Find_Error

    Text = UCase$(Text)

    For i = 1 To Message.Count
        If Message.Item(i).Text = Text Then
            FindItem = i
            Exit Function
        End If
    Next

    Exit Function

Find_Error:
    FindItem = 0
End Function

' Find the key - return the index (0 = not found)
Public Function FindKey(ByVal Key As String) As Boolean
Attribute FindKey.VB_Description = "Finds an item by item key string. Returns the index. Zero when not found."
    Dim sDummy As String

    On Local Error GoTo Find_Error

    sDummy = Message.Item(Key).Text
    FindKey = True
    Exit Function

Find_Error:
    FindKey = False
End Function

' Remove selected message
Public Sub RemoveItem(Index As Variant)
Attribute RemoveItem.VB_Description = "Removes an item to the message queue. Index can be the index number or the key string."
    On Local Error GoTo Remove_Error

    If ValidIndexType(Index) Then
        Message.Remove Index
        SetControlState
    End If

Remove_Error:
End Sub

' Get the count
Public Function ListCount() As Long
    ListCount = Message.Count
End Function

' Remove all messages
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all messages in the queue"
    Do While Message.Count > 0
        Message.Remove 1
    Loop
    SetControlState
End Sub

Public Property Get ItemData(Index As Variant) As Variant
Attribute ItemData.VB_Description = "Additional item data referenced by index or key. Contains a variant value."
Attribute ItemData.VB_ProcData.VB_Invoke_Property = ";Data"
    If ValidIndexType(Index) Then
        On Local Error Resume Next
        ItemData = Message.Item(Index).ItemData
    End If
End Property
Public Property Let ItemData(Index As Variant, ByVal vNewValue As Variant)
    If ValidIndexType(Index) Then
        On Local Error Resume Next
        Message.Item(Index).ItemData = vNewValue
    End If
End Property

Public Property Get ItemColor(Index As Variant) As Long
    If ValidIndexType(Index) Then
        On Local Error Resume Next
        ItemColor = Message.Item(Index).Colour
    End If
End Property
Public Property Let ItemColor(Index As Variant, ByVal vNewValue As Long)
    If ValidIndexType(Index) Then
        On Local Error Resume Next
        Message.Item(Index).Colour = vNewValue
        If mPan.Moving Then PaintControl
    End If
End Property

Public Property Get List(Index As Variant) As String
Attribute List.VB_Description = "Item list referenced by index or key."
Attribute List.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute List.VB_UserMemId = 0
    If ValidIndexType(Index) Then
        On Local Error Resume Next
        List = Message.Item(Index).Text
    End If
End Property
Public Property Let List(Index As Variant, ByVal vNewValue As String)
    On Local Error GoTo List_Error
    If ValidIndexType(Index) Then
        On Local Error Resume Next
        Message.Item(Index).Text = vNewValue
        SetControlState
        If mPan.Moving Then PaintControl
    End If
List_Error:
End Property
