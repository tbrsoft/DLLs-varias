Attribute VB_Name = "Gobales"

'para mover una ventana sabiendo su HWND
Public Declare Function MoveWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal X As Long, _
    ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'para saber la posicion y tamaño de _
    una ventana sabiendo su HWND
Public Declare Function GetWindowRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
