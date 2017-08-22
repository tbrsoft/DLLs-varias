VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private colTags As Collection

Public Property Get TagCount() As Long
    TagCount = colTags.Count
End Property

Friend Sub AddTag( _
    ByVal name As String, _
    ByVal value As String _
)

    Dim clsTag  As New StreamTag

    clsTag.SetValues name, value
    colTags.Add clsTag
End Sub

Friend Sub ClearTags()
    Dim i   As Long

    For i = colTags.Count To 1 Step -1
        colTags.Remove i
    Next
End Sub

Public Property Get TagItem( _
    ByVal index As Long _
) As StreamTag

    Set TagItem = colTags(index)
End Property

Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = colTags.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set colTags = New Collection
End Sub
