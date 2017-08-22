Attribute VB_Name = "Globales"
Public tbrPintaNoPix As tbrPintar

Public Type RECT
  qLeft As Long
  qTop As Long
  qRight As Long
  qBottom As Long
End Type

Public terr As New tbrErrores.clsTbrERR
Public fso As New Scripting.FileSystemObject

Public qAlgunFormulario As Form
Public tmpFolder As String 'carpeta donde ir poniendo archivos temporales

'lo necesito en tantos lugares  ...
Public HechoParaPixlesAncho As Long
Public HechoParaPixlesAlto As Long
'se refiere a para que tamaño de pantalla fue hecho
'lo leo y grabo como propiedad del skin, esto podria traer problemas si reutilizo objetos
'sueltos que fueron hecho para funcionar en un tamaño y despues se usan en un skin con
'otras especificaciones. SEGUIRAQUI. no es un problema grabe pero en el futuro se podria
'hacer alguna adaptacion al agregar objetos externos al skin
'mientras tanto el valor predeterminado es 1024 x 768


Public Sub Main()
    
End Sub



