VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFullImgs_FILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cada uno de los archivos que forman parte del skin

Public Enum TypeFileSkin
    mTextFile = 1
    mImageJPG = 2
    mImagePNG = 3
    mImageBMP = 4
    mImageGIF = 5
    mBinary = 6
End Enum

Private mBaseName As String 'solo lectura nombre del archivo sin path ni extencion
Private mPath As String
Private mType As TypeFileSkin
Private mID As String 'identificador unico. A veces los skins guardan los archivos con diferente nombre

Public Property Get path() As String
    path = mPath
End Property

Public Property Let path(newV As String)
    mPath = newV
    mBaseName = fso.GetBaseName(mPath)
    Select Case LCase(fso.GetExtensionName(mPath))
        Case "jpg", "jpeg"
            mType = mImageJPG
        
        Case "bmp"
            mType = mImageBMP
        
        Case "gif"
            mType = mImageGIF
            
        Case "png"
            mType = mImagePNG
            
        Case "tif", "tiff" 'OTROS FORMATOS DE IMAGENES COMPATIBLES
            mType = mImageJPG
            
        Case "txt", "log"
            mType = mTextFile
        
        Case "dat"
            mType = mBinary
    
    End Select
End Property

Public Property Get BaseName() As String
    BaseName = mBaseName
End Property

Public Property Get TypeFile() As TypeFileSkin
    TypeFile = mType
End Property

Public Property Let TypeFile(newV As TypeFileSkin)
    mType = newV
End Property

Private Function GetStringType(typ As TypeFileSkin) As String
    Select Case typ
        Case 1: GetStringType = "Texto"
        Case 2: GetStringType = "Imagen Jpg"
        Case 3: GetStringType = "Imagen Png"
        Case 4: GetStringType = "Imagen Bmp"
        Case 5: GetStringType = "Imagen Gif"
        Case 6: GetStringType = "Binary"
        Case Else: GetStringType = "Desconocido(" + CStr(typ) + ")"
    End Select
End Function

Public Function toString() As String
    toString = "SkinFile:" + fso.GetBaseName(mPath) + vbCrLf + "  Path:" + mPath + vbCrLf + "  Type:" + GetStringType(mType)
End Function

Public Property Get sID() As String
    sID = mID
End Property

Public Property Let sID(newV As String)
    mID = newV
End Property
