VERSION 5.00
Begin VB.Form frmDXCallback 
   Caption         =   "Form1"
   ClientHeight    =   180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   ScaleHeight     =   180
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDXCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' DirectX Events can only be created in Forms

Implements DirectXEvent8

Private Sub DirectXEvent8_DXCallback( _
    ByVal eventid As Long _
)

    modEventManager.GetEventClass(eventid).OnEvent eventid
End Sub

Public Function CreateEvent() As Long
    CreateEvent = DirectX.CreateEvent(Me)
End Function
