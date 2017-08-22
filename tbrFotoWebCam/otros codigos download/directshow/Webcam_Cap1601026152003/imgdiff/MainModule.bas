Attribute VB_Name = "MainModule"
Option Explicit
'
Public Const COLOR_BASE = 16777216
'
Public iTolerance As Integer
Public bFastScan As Boolean
Public iScanMethod As Integer
'
Public pTmpPic As PictureBox

Public Type RGBpoint
  Red As Byte
  Green As Byte
  Blue As Byte
End Type

Public Type RGBthingy
  Value As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
       (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

