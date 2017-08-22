<HTML>
<HEAD>
<TITLE>Broadcast WebClient Demo</TITLE>
</HEAD>
<BODY>
<OBJECT ID="VideoWindow"
CLASSID="CLSID:9AA6E79E-A911-47BB-B4B0-63D2565D38F5"
CODEBASE="WebClient.CAB#version=1,0,0,0">
</OBJECT>
<SCRIPT LANGUAGE="VBSCRIPT">
<%@ Language = "VBScript" %>
<% Response.Write "VideoWindow.SetHost """ & Request.ServerVariables("LOCAL_ADDR") & """" & vbCrLf %>
<% Response.Write "VideoWindow.PlayFile """ & "FINDFILE.AVI" & """" %>
</SCRIPT>
</BODY>
</HTML>