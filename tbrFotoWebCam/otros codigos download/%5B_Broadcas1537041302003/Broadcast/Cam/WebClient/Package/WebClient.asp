<HTML>
<HEAD>
<TITLE>Broadcast WebClient Demo</TITLE>
</HEAD>
<BODY>
<OBJECT ID="VideoWindow"
CLASSID="CLSID:ED60998B-72E2-40F0-8AC5-AF458740C39D"
CODEBASE="WebClient.CAB#version=1,0,0,0">
</OBJECT>
<SCRIPT LANGUAGE="VBSCRIPT">
<%@ Language = "VBScript" %>
<% Response.Write "VideoWindow.Connect """ & Request.ServerVariables("LOCAL_ADDR") & """" %>
</SCRIPT>
</BODY>
</HTML>

