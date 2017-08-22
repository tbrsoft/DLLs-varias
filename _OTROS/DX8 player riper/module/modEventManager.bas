Attribute VB_Name = "modEventManager"
Option Explicit

' DirectX Event Reciever Manager
'
' frmDXCallback recieves events,
' modEventManager assigns them to
' the interface which has created them

Private colEvents   As Collection

Public Sub InitEventManager()
    If colEvents Is Nothing Then
        Set colEvents = New Collection
    End If
End Sub

Public Sub AddEvent( _
    cb As IDXCallback _
)

    colEvents.Add cb
End Sub

Public Sub RemEvent( _
    ByVal eventid As Long _
)

    On Error Resume Next

    Dim i       As Long
    Dim arr()   As Long

    For i = 1 To colEvents.Count
        arr = colEvents.Item(i).MyEvents
        If InI4Array(arr, eventid) Then
            colEvents.Remove i
        End If
    Next
End Sub

Public Function GetEventClass( _
    ByVal eventid As Long _
) As IDXCallback

    Dim idx As IDXCallback

    For Each idx In colEvents
        If InI4Array(idx.MyEvents, eventid) Then
            Set GetEventClass = idx
            Exit Function
        End If
    Next
End Function

Private Function InI4Array( _
    arr() As Long, _
    ByVal value As Long _
) As Boolean

    Dim i   As Long

    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            InI4Array = True
            Exit Function
        End If
    Next
End Function
