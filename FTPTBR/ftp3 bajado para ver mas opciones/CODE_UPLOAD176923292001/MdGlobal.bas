Attribute VB_Name = "MdGlobal"
Public inigo As cIniFile
Private Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type

Declare Function CreateFile Lib "kernel32" _
Alias "CreateFileA" _
(ByVal lpFileName As String, _
ByVal dwDesiredAccess As Long, _
ByVal dwShareMode As Long, _
lpSecurityAttributes As SECURITY_ATTRIBUTES, _
ByVal dwCreationDisposition As Long, _
ByVal dwFlagsAndAttributes As Long, _
ByVal hTemplateFile As Long) As Long

Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const CREATE_NEW = 1
Private Const CREATE_ALWAYS = 2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Dim h1 As Long

Function CreateIt(file As String)
Dim sec As SECURITY_ATTRIBUTES
sec.bInheritHandle = True
sec.lpSecurityDescriptor = 0
sec.nLength = Len(sec)

  h1 = CreateFile(file, GENERIC_READ, 0, sec, CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, 0)
                  
                  CloseHandle (h1)

End Function
Function Validate_File(ByVal Filename As String) As Boolean

       Dim fileFile As Integer
       '     'attempt to open file
       fileFile = FreeFile
       On Error Resume Next
       Open Filename For Input As fileFile
       '     'check for error

              If Err Then
                     Validate_File = False
              Else
                     '     'file exists
                     '     'close file
                     Close fileFile
                     Validate_File = True
              End If

End Function
