VERSION 5.00
Begin VB.UserControl ctlEBSlider 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   ScaleHeight     =   360
   ScaleWidth      =   4440
   ToolboxBitmap   =   "ctlEBSlider.ctx":0000
   Begin VB.PictureBox picSlider 
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Line linGroove 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   4380
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line linGroove 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   4380
      Y1              =   180
      Y2              =   180
   End
End
Attribute VB_Name = "ctlEBSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'[Description]
'   EBSlider
'   A stand-alone slider control

'[Author]
'   Richard Allsebrook  <RA>    RichardAllsebrook@earlybirdmarketing.com

'[History]
'   V1.0.0  20/06/2001
'   Initial Release

'[Declarations]

'Property storage
Private lngMin              As Long         'Minimum value range
Private lngMax              As Long         'Maximum value range
Private lngValue            As Long         'Current Value
Private lngSliderWidth      As Long
Private zBorderStyle        As EBSliderBorderStyle

Private zOrientation        As EBSliderOrientation  'Current Orientation

'Event Stubs
Event Changed()

'Enums
Public Enum EBSliderOrientation
    EBHorizontal
    EBVertical
End Enum

Public Enum EBSliderBorderStyle
    EBNone = 0
    EBSunkenOuter = &H2
    EBRaisedInner = &H4
    EBEtched = (EBSunkenOuter Or EBRaisedInner)
End Enum

'API Stubs
Private Declare Function DrawEdge Lib "user32" ( _
  ByVal hdc As Long, _
  qrc As RECT, _
  ByVal edge As Long, _
  ByVal grfFlags As Long _
     ) As Long
     
Private Declare Function SetRect Lib "user32" ( _
  lpRect As RECT, _
  ByVal X1 As Long, _
  ByVal Y1 As Long, _
  ByVal X2 As Long, _
  ByVal Y2 As Long _
    ) As Long

'API UDTs
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'API Constants
Private Const BDR_RAISEDINNER = &H4

Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Sub picSlider_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'[Description]
'   Allow the user to reposition the slider by dragging

'[Declarations]
Dim lngPos                  As Long         'New position of slider
Dim sglScale                As Single       'Calculated scale of slider

'[Code]

    If Button = vbLeftButton Then
        'Only move if the button is pressed
    
        With picSlider
        
            If zOrientation = EBHorizontal Then
                'calulate new position of slider and round to nearest pixel
                lngPos = ((.Left + x - lngSliderWidth / 2) \ 15) * 15
                    
                'Constrain to control
                If lngPos < 0 Then
                    'Attempted to move slider past start
                    lngPos = 0
                
                ElseIf lngPos > UserControl.Width - lngSliderWidth Then
                    'Attempted to move slider past end
                    lngPos = UserControl.Width - lngSliderWidth
                End If
                
                'Move slider
                .Left = lngPos
                
                'Re-calculate value based on new position
                sglScale = (UserControl.Width - lngSliderWidth) / (lngMax - lngMin)
                lngValue = (lngPos / sglScale) + lngMin
                
                RaiseEvent Changed
            
            Else
        
            'Vertical
                'calulate new position of slider and round to nearest pixel
                lngPos = ((.Top + Y - lngSliderWidth / 2) \ 15) * 15
                    
                'Constrain to control
                If lngPos < 0 Then
                    'Attempted to move slider past start
                    lngPos = 0
                
                ElseIf lngPos > UserControl.Height - lngSliderWidth Then
                    'Attempted to move slider past end
                    lngPos = UserControl.Height - lngSliderWidth
                End If
                
                'Move slider
                .Top = lngPos
                
                'Re-calculate value based on new position
                sglScale = (UserControl.Height - lngSliderWidth) / (lngMax - lngMin)
                lngValue = (lngPos / sglScale) + lngMin
                
                RaiseEvent Changed
            End If
            
        End With
        
    End If

End Sub

Private Sub picSlider_Paint()

'[Description]
'   Draw a raised border round the slider

'[Declarations]
Dim udtRECT                 As RECT         'Slider RECT structure

'[Code]

    With picSlider
        SetRect udtRECT, 0, 0, .Width / 15, .Height / 15
        DrawEdge .hdc, udtRECT, BDR_RAISEDINNER, BF_RECT
    End With
    
End Sub

Private Sub picSlider_Resize()

    picSlider.Cls
    
End Sub

Private Sub UserControl_InitProperties()

'[Description]
'   Set initial values for properties

'[Code]

    lngMin = 0
    lngMax = 100
    lngValue = 50
    lngSliderWidth = 315
    picSlider.BackColor = vb3DFace
    Orientation = EBHorizontal
    BorderStyle = EBNone
        
    'Initialise the slider
    PositionSlider
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'[Description]
'   Clicking anywhere on the control makes the slider jump to that position

'[Declarations]
Dim lngPos                  As Long         'New position of slider
Dim sglScale                As Single       'Calculated scale of slider

    
    With picSlider
    
        If zOrientation = EBHorizontal Then
            'Caluclate new position and round to nearest pixel
            lngPos = ((x - lngSliderWidth / 2) \ 15) * 15
            
            'Constrain to control
            If lngPos < 0 Then
                'Attempted to move past start
                lngPos = 0
            
            ElseIf lngPos > UserControl.Width - lngSliderWidth Then
                'Attempted to move past end
                lngPos = UserControl.Width - lngSliderWidth
            End If
            
            'Move slider
            .Left = lngPos
            
            'Calculate value based on new position
            sglScale = (UserControl.Width - .Width) / (lngMax - lngMin)
            lngValue = (lngPos / sglScale) + lngMin
            
            RaiseEvent Changed
        Else
            'Caluclate new position and round to nearest pixel
            lngPos = ((Y - lngSliderWidth / 2) \ 15) * 15
            
            'Constrain to control
            If lngPos < 0 Then
                'Attempted to move past start
                lngPos = 0
            
            ElseIf lngPos > UserControl.Height - lngSliderWidth Then
                'Attempted to move past end
                lngPos = UserControl.Height - lngSliderWidth
            End If
            
            'Move slider
            .Top = lngPos
            
            'Calculate value based on new position
            sglScale = (UserControl.Height - lngSliderWidth) / (lngMax - lngMin)
            lngValue = (lngPos / sglScale) + lngMin
            
            RaiseEvent Changed
        
        End If
        
    End With
            
End Sub

Private Sub UserControl_Paint()

Dim udtRECT                 As RECT
    
    SetRect udtRECT, 0, 0, UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY
    DrawEdge UserControl.hdc, udtRECT, zBorderStyle, BF_RECT
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'[Description]
'   Retrieve stored properties from PropBag

'[Code]

    With PropBag
        lngMin = .ReadProperty("Min", 0)
        lngMax = .ReadProperty("Max", 100)
        lngValue = .ReadProperty("Value", 50)
        lngSliderWidth = .ReadProperty("SliderWidth", 315)
        picSlider.BackColor = .ReadProperty("SliderColor", vb3DFace)
        BorderStyle = .ReadProperty("BorderStyle", EBNone)
        Orientation = .ReadProperty("Orientation", EBHorizontal)
    End With
        
    'Initialise the slider
    PositionSlider
    
End Sub

Private Sub UserControl_Resize()

'[Description]
'   Resize constituant controls to match new control size

'[Declarations]
Dim lngWidth                As Long             'New control width
Dim lngHeight               As Long             'New control height

Dim intIndex                As Integer
'[Code]

    With UserControl
        .Cls
    
        lngWidth = .Width - Screen.TwipsPerPixelX
        lngHeight = .Height - Screen.TwipsPerPixelY
        
        If zOrientation = EBHorizontal Then
            'Horizontal
            
            For intIndex = 0 To 1
                linGroove(intIndex).X1 = 15
                linGroove(intIndex).X2 = lngWidth - 15
                linGroove(intIndex).Y1 = lngHeight / 2
                linGroove(intIndex).Y2 = lngHeight / 2
            Next
            
            picSlider.Top = 0
            picSlider.Height = lngHeight
            picSlider.Width = lngSliderWidth
            
        Else
            'Vertical
            
            For intIndex = 0 To 1
                linGroove(intIndex).X1 = lngWidth / 2
                linGroove(intIndex).X2 = lngWidth / 2
                linGroove(intIndex).Y1 = 15
                linGroove(intIndex).Y2 = lngHeight - 15
            Next
                        
            picSlider.Left = 0
            picSlider.Width = lngWidth
            picSlider.Height = lngSliderWidth
        End If
        
    End With
    
    'Initialise the slider
    PositionSlider
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'[Description]
'   Store properties in PropBag

'[Code]

    With PropBag
        .WriteProperty "Min", lngMin, 0
        .WriteProperty "Max", lngMax, 100
        .WriteProperty "Value", lngValue, 50
        .WriteProperty "SliderWidth", lngSliderWidth, 315
        .WriteProperty "SliderColor", picSlider.BackColor, vb3DFace
        .WriteProperty "BorderStyle", zBorderStyle, EBNone
        .WriteProperty "Orientation", zOrientation, EBHorizontal
    End With
    
End Sub

Private Function PositionSlider()

'[Description]
'   Moves the slider to match the current Value property

'[Declarations]
Dim sglScale                As Single       'Calculated scale of slider

'[Code]

    With picSlider
    
        If lngMax - lngMin <> 0 Then
            'Avoid devide by zero error
            
            'Calculate new position
            
            If zOrientation = EBHorizontal Then
                sglScale = (UserControl.Width - lngSliderWidth) / (lngMax - lngMin)
                .Left = (lngValue - lngMin) * sglScale
            Else
                sglScale = (UserControl.Height - lngSliderWidth) / (lngMax - lngMin)
                .Top = (lngValue - lngMin) * sglScale
            End If
            
        End If
        
    End With
    
End Function

Public Property Get Min() As Long

'[Description]
'   Return the current Min property

'[Code]

    Min = lngMin
    
End Property

Public Property Let Min(NewValue As Long)

'[Description]
'   Set the Min property

'[Code]

    If NewValue <= lngMax Then
        'Min must be less than Max
        lngMin = NewValue
        
        If lngValue < lngMin Then
            'ensure current value still in min-max range
            lngValue = lngMin
            PropertyChanged "Value"
        End If
        
        PositionSlider
        
        PropertyChanged "Min"
    End If
    
End Property

Public Property Get Max() As Long

'[Description]
'   Return the current Max property

'[Code]

    Max = lngMax
    
End Property

Public Property Let Max(NewValue As Long)

'[Description]
'   Set the current max property

'[Code]

    If NewValue > lngMin Then
        'Max must be greater than Min
        lngMax = NewValue
        
        If lngValue > lngMax Then
            'Ensure current value is within new min-max range
            lngValue = lngMax
            PropertyChanged "Value"
        End If
        
        'Re-initialise slider
        PositionSlider
        
        PropertyChanged "Max"
    End If
    
End Property

Public Property Get Value() As Long

'[Description]
'   Return the current Value property

'[Code]

    If zOrientation = EBHorizontal Then
        Value = lngValue
    Else
        Value = lngMax + lngMin - lngValue
    End If
    
End Property

Public Property Let Value(NewValue As Long)

'[Description]
'   Set the current Value property

'[Code]

    'Constrain new value to min-max range
    If NewValue < lngMin Then
        NewValue = lngMin
    
    ElseIf NewValue > lngMax Then
        NewValue = lngMax
    End If
    
    lngValue = NewValue
    
    'Reposition slider
    PositionSlider
    
    PropertyChanged "Value"
    RaiseEvent Changed
    
End Property

Public Property Get SliderWidth() As Long

'[Description]
'   Reurn current slider width

'[Code]

    SliderWidth = lngSliderWidth
    
End Property

Public Property Let SliderWidth(NewValue As Long)

'[Description]
'   Set slider width

'[Code]

    If (zOrientation = EBHorizontal And NewValue < UserControl.Width) _
      Or (zOrientation = EBVertical And NewValue < UserControl.Height) Then
        'Ensure slider width is less than control
        lngSliderWidth = NewValue
        
        If zOrientation = EBHorizontal Then
            picSlider.Width = lngSliderWidth
            picSlider.Height = UserControl.Height
        Else
            picSlider.Height = lngSliderWidth
            picSlider.Width = UserControl.Width
        End If
        
        'Redraw the slider
        picSlider_Paint
        
        'Reposition the slider
        PositionSlider
        
        PropertyChanged "SliderWidth"
    End If
    
End Property

Public Property Get SliderColor() As OLE_COLOR

'[Description]
'   Return the current slider color

'[Code]

    SliderColor = picSlider.BackColor
    
End Property

Public Property Let SliderColor(NewValue As OLE_COLOR)

'[Description]
'   Set the slider color

'[Code]

    picSlider.BackColor = NewValue
    
    'Redraw the slider
    picSlider_Paint
    
    PropertyChanged "SliderColor"
    
End Property

Public Property Get Orientation() As EBSliderOrientation

    Orientation = zOrientation
    
End Property

Public Property Let Orientation(NewValue As EBSliderOrientation)

    zOrientation = NewValue
    SliderWidth = lngSliderWidth 'force resize or slider
    
    picSlider_Paint
    UserControl_Resize
    
End Property

Public Property Get BorderStyle() As EBSliderBorderStyle

    BorderStyle = zBorderStyle
    
End Property

Public Property Let BorderStyle(NewValue As EBSliderBorderStyle)

Dim udtRECT                 As RECT

    zBorderStyle = NewValue
    UserControl_Paint
    
End Property
