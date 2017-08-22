Attribute VB_Name = "modMain"
Global doing_multi As Boolean
Global Doing_Download As Boolean
Global Doing_Upload As Boolean
Global Response As Integer
Global Retrned As Long
Global remote_dir As String
Global local_dir As String
Global halt_transfer As Boolean
Global servu As Boolean
Global ServPort As Integer
Global r_Count As String
Global inigo As cIniFile
Global SizeIt As Long
Global CurrDNFile As String
Global UserName As String
Global Password As String

Global fName() As String
Global fPath() As String

Type Com
       BackCode As String
       Command As String
End Type

Type ClientInfo
       File_Name As String
       full_count As Integer
       total_size As Long
       fFile As Long
       currentFile As String
       transferTotalBytes As Long
       transferBytesSent As Long
End Type


Public client As ClientInfo


Global Buffer As Long
Global States(3) As Com
Global state As Integer
Public Sub open_file(filename As String)
On Error GoTo here

    If Right(local_dir, 1) <> "\" Then
    local_dir = local_dir & "\"
    End If
    
    If Left(local_dir, 1) = "\" Then
    local_dir = Mid(local_dir, 2)
    End If
    
    Open (local_dir & filename) For Binary Access Write As #1

here:
DoEvents
End Sub

Function Parse2Array(ByVal strText As String, ByRef strArray() As String, ByVal strDelim As String) As Long
       Dim intPos As Long
       Dim intIndex As Long
       strText = Trim(strText)
       ReDim strArray(10) As String
       Do While strText <> ""
           If intIndex > UBound(strArray()) Then
               ReDim Preserve strArray(intIndex + 20)
           End If
           intPos = InStr(1, strText, strDelim)
           If intPos > 0 Then
               strArray(intIndex) = Left(strText, InStr(1, strText, strDelim) - 1)
               strText = Trim(Mid(strText, InStr(1, strText, strDelim) + 1))
           Else
               strArray(intIndex) = strText
               Exit Do
           End If
           intIndex = intIndex + 1
       Loop
       ReDim Preserve strArray(intIndex) As String
       Parse2Array = UBound(strArray())
   End Function
Function CountStr(ByVal parseStringx, Parser As String) As Variant
On Error Resume Next
Dim lastPos As Integer
Dim subPos As Integer
Dim argPos(1 To 500) As Integer
Dim argContent(1 To 500)
parsestring = parseStringx
parsestring = Trim(Right(parsestring, ((Len(parsestring)) - (InStr(parsestring, Parser)))))

parsestring = parsestring & vbCrLf
argcount = 0
Do
    DoEvents
    lastPos = InStr((lastPos + 1), parsestring, vbCrLf)
    If lastPos = 0 Then Exit Do
    argcount = argcount + 1
    argPos(argcount) = lastPos
Loop
If argcount = 0 Then Exit Function
CountStr = argcount
Exit Function
End Function
Function DoTheIni(Up As Boolean) As Boolean
Dim ret As Boolean
Dim o As Integer, c As Integer
Dim direct As String

   DoTheIni = False
   
   If Up = True Then
   direct = "AutoUP"
   ElseIf Up = False Then
   direct = "AutoDN"
   End If
   
   c = 0

   inigo.Section = direct
   inigo.Key = "count"
   r_Count = inigo.Value
   inigo.Key = "ChDirName"
   remote_dir = inigo.Value

If r_Count = "" Then
   r_Count = 0
   DoTheIni = False
Exit Function
End If

For o = 1 To r_Count
  inigo.Key = "file" & o
  
  If Up = False Then GoTo down
  
  ret = Validate_File(inigo.Value)
  
  If ret = False Then GoTo missed
down:
  
  c = c + 1
            ReDim Preserve fName(1 To c)
               fName(c) = inigo.Value
            
            Open fName(c) For Binary As #4
               SizeIt = SizeIt + LOF(4)
            Close 4
             
             If Up = True Then
  inigo.Key = "path" & o
            ReDim Preserve fPath(1 To c)
            fPath(c) = inigo.Value
             End If
            
missed:
DoEvents
Next

   DoTheIni = True
End Function
Public Function ExtractName(chrsin As String)
On Error Resume Next
If InStr(chrsin, "\") Then 'check to see if a forward slash exists
   For idx = Len(chrsin) To 1 Step -1 'step though until full name is extracted
       If Mid(chrsin, idx, 1) = "\" Then
          chrsout = Mid(chrsin, idx + 1)
          Exit For
       End If
   Next idx
ElseIf InStr(chrsin, ":") = 2 Then 'otherwise, check to see if a colon exists
   chrsout = Mid(chrsin, 3)        'if so, return the filename
Else
   chrsout = chrsin 'otherwise, return the original string
End If
     
ExtractName = chrsout 'return the filename to the user
End Function
Function Validate_File(ByVal filename As String) As Boolean

       Dim fileFile As Integer
       '     'attempt to open file
       fileFile = FreeFile
       On Error Resume Next
       Open filename For Input As fileFile
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

