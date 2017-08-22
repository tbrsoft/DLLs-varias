Attribute VB_Name = "Mod1"
Option Explicit

Public Function AppPath(ByVal zPath As String) As String
  If Right$(zPath, 1) = "\" Then AppPath = zPath Else AppPath = zPath & "\"
End Function
