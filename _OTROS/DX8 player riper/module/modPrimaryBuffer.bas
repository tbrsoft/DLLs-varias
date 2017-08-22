Attribute VB_Name = "modPrimaryBuffer"
Option Explicit

' initialize DirectX

Public DirectX          As DirectX8

Public Function InitializeDirectX() As Boolean
    On Error GoTo ErrorHandler

    If DirectX Is Nothing Then
        Set DirectX = New DirectX8
    End If

    InitializeDirectX = True

ErrorHandler:
End Function
