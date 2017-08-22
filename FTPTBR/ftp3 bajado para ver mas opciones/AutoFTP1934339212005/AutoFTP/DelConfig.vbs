Option Explicit

Dim objShell, objSysdr, objSysdrItem, objFSO, Sys32
Const System32 = &H25&

'most of this code is for retrieving the path to the system32 folder where the configuration file is stored.

Set objShell = CreateObject("Shell.Application")
Set objSysdr = objShell.Namespace(System32)
Set objSysdrItem = objsysdr.Self
Sys32 = objsysdrItem.Path
Set objFSO = CreateObject("Scripting.FileSystemObject")

'we delete the configuration file if it exists and generate a message to inform the user.

If objFSO.FileExists(Sys32 & "\ftp32") Then
	objFSO.DeleteFile(Sys32 & "\ftp32")
	msgbox "The AutoFTPScript configuration file is now deleted!   ", 64, "File removed"
else
	msgbox "The AutoFTPScript configuration file does not exist!   " & vbcrlf & vbcrlf & "Run AutoFTPScript to create the FTP upload" & vbcrlf & "settings configuration file.", 48, "File not found"
end if