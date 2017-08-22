Option Explicit

Dim objSysdr, objSysdrItem , Sys32, objFSO, objMyFile, objShell, Ts
Dim strFTPServer, strLoginID, strPassword, strFTPServerFolder, strFTPDatFileName, strFileToPut, strUplMode, strIsFTPpath, strIsNoFTPpath 
Dim blFileCount, StartMess

Const System32 = &H25&
Set objShell = CreateObject("Shell.Application")
Set objSysdr = objShell.Namespace(SYSTEM32)
Set objSysdrItem = objsysDR.Self
Sys32 = objsysDRItem.Path
strFTPDatFileName = Sys32 & "/ftp32"

Set objFSO = CreateObject("Scripting.FileSystemObject")

' the largest part of this script is the configuration part. You will get about 7 inputboxes the first time
' you run this script (and also after using DelConfig.vbs) in which you can input your FTP settings.
' The next time the script is run, it will upload the files you specified during the configuration.

   If not objFSO.FileExists(strFTPDatFileName ) Then

	StartMess = MsgBox (vbcrlf & vbcrlf & "ATTENTION" & vbcrlf & "=======" & vbcrlf & vbcrlf & "Because this is the first time you are using this script you have to enter some configuration parameters.          " & vbcrlf & vbcrlf & "You have to do this only once. If you would like to change the configuration later, (FTP server data" & vbcrlf & "or upload file names) then you must run the in this package included DelConfig vbscript first to delete" & vbcrlf & "the already existing configuration data." & vbcrlf & vbcrlf & vbcrlf & vbcrlf & vbcrlf & vbcrlf & "© Kurt Koenig * Belgium  09/2005" & vbcrlf & vbcrlf, 48, "Initializing AutoFTPScript configuration ...")

	strFtpServer = InputBox("Step 1 of 7" & vbcrlf & vbcrlf & "Enter the name of your FTP host without ftp://" & vbcrlf & "Example: corporate.skynet.be", "Your FTP host name ...","", 5000, 5500)
	strLoginID = InputBox("Step 2 of 7" & vbcrlf & vbcrlf & "Enter your FTP login name ...", "Your user name ...","", 5000, 5500)
	strPassword  =  InputBox("Step 3 of 7" & vbcrlf & vbcrlf & "Enter your FTP password ...", "Your password ...","", 5000, 5500)
	strFTPServerFolder = InputBox("Step 4 of 7" & vbcrlf & vbcrlf & "Your FTP subdirectory after the slash ..." & vbcrlf  & "Just leave the forward slash in case of no sub directory (/)", "Your subdirectory...","/", 5000, 5500)
	strIsFTPpath = InStr(1,strFTPServerFolder,"/",1)
	strIsNoFTPpath = InStr(1,strFTPServerFolder,"\",1)

' in case of an accidental use of a back slash instead of a forward slash, a messagebox is generated and the script will quit.


		if strIsNoFTPpath then
			msgbox "You can't use a back slash in the FTP folder path, use forward slashes instead." & vbcrlf & vbcrlf & "Start this configuration again please!"
			WScript.quit
		end if

' if the first character isn't a forward slash, a messagebox is generated and the script will also quit.

		if not strIsFTPpath = 1 then
			msgbox "The first character must be a forward slash!" & vbcrlf & vbcrlf & "Start this configuration again please!"
			WScript.quit
		end if		

	strUplMode  =   InputBox("Step 5 of 7" & vbcrlf & vbcrlf & "Enter ascii for text only, or bin (binary) " & vbcrlf  & "for all kind of other files.", "Uploadmodus","bin", 5000, 5500)
	blFileCount = InputBox("Step 6 of 7" & vbcrlf & vbcrlf & "Would you like to upload one or multiple files?" & vbcrlf & "Enter single for a single file, or enter multi for multiple files.", "Put or multiple put.","multi", 5000, 5500)
	

' when using anything else then the words "single" or "multi" in this input, a messagebox is generated and again the script will quit.
' trim is used to remove trailing or preceding spaces if there are any (I didn't test if this is necessary, just to make sure)

		if not trim(blFileCount) = "single" and not trim(blFileCount) = "multi" then
			msgbox "You can't use other values then single or multi!" & vbcrlf & vbcrlf & "Start this configuration again please!"
			WScript.quit
		end if

	strFileToPut = InputBox("Step 7 of 7" & vbcrlf & vbcrlf & "The path to the file(s) to be uploaded" & vbcrlf  & "Use forward slashes in the path, wildcards are also allowed (Example C:/FTPfolder/*.*)", "Local upload path","C:/*", 5000, 5500)
	strIsNoFTPpath = InStr(strFileToPut,"\")


' again, if you used a back slash instead of a forward slash, a messagebox is generated and the script will quit.

		if strIsNoFTPpath then
			msgbox "You can't use a back slash in the file path, use forward slashes instead." & vbcrlf & vbcrlf &"Start this configuration again please!"
			WScript.quit
		end if

' here we write the settings to a file, if the file exists this script will skip 
' the configuration part the next time and go directly to the last "else" in this script.


	Set objMyFile = objFSO.CreateTextFile(strFTPDatFileName, True)
	objMyFile.WriteLine ("open " & strFTPServer)
	objMyFile.WriteLine (strLoginID)
	objMyFile.WriteLine (strPassword)
	objMyFile.WriteLine ("cd " & strFTPServerFolder)
	objMyFile.WriteLine (strUplMode)

	    
		if blFileCount = "multi" then
			objMyFile.WriteLine ("mput " & strFileToPut)
		else
			objMyFile.WriteLine ("put " & strFileToPut)
		end if
	  

	objMyFile.WriteLine ("bye")
	objMyFile.Close
	Set objMyFile = Nothing
	StartMess = MsgBox (vbcrlf & vbcrlf & "DONE!" & vbcrlf & "====" & vbcrlf & vbcrlf & "Your FTP configuration data is saved." & vbcrlf & "Uploading will start the next time you run this script.       "& vbcrlf & vbcrlf & vbcrlf & vbcrlf & vbcrlf & vbcrlf & "© Kurt Koenig * Belgium  09/2005" & vbcrlf & vbcrlf, 48, "End of AutoFTPScript configuration")
   else

	Set objShell = WScript.CreateObject( "WScript.Shell")
	objShell.Run ("ftp -i -s:" & chr(34) & strFTPDatFileName & chr(34))

   end if

Set objFSO = Nothing
Set objMyFile = Nothing
Set objShell = Nothing