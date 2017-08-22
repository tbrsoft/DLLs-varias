wcamap for Windows - Source code for Visual Basic 6

A simple video capture application using the DirectShow technology.

New in V2 : a simple motion detection ( difference between two images ).

You may freely use and modify the source code contained in this product.  You may also freely distribute any application that uses this sample code or derivations thereof.  However, you may not redistribute 
any part of this archive or in any way derive financial gain from this sample without the express permission of it's authors.


Comments, suggestions, and bug reports can be sent to mlinari@hotmail.com

IMPORTANT. THIS PROGRAM NEED THE FOLLOWING DLLs IN ORDER TO WORK.
=================================================================


- CapStill.dll, FSFWrap.dll, quartz.dll 

These are all dll for the use of DirectShow in VB. You can download it for free on:

http://www.gdcl.co.uk/index.htm

- converter.dll

This is for converting BMP to JPG. You can download it for free on :

http://www.visual-basic.it/download.asp

( The package name is bmptojpg.zip )

You must also download the "ntsvc.ocx v1.1" and "systray.ocx v1" if you want use this program as service ( you can found it quite everywhere ).

========

The project is dived in two parts :  the wcam.zip and the imgdiff.zip. YOU MUST FIRST DECOMPRESS AND COMPILE THE IMGDIFF.ZIP IN ORDER TO INSTALL THE IMGDIFF.DLL FOR THE "MOTION DETECTION"

========


You can run this program as a normal application or use it as a service ( if you have Windows NT/2000/XP ).
 
 - Installation of the service : wcamcap -install
 - Uninstall the service : wcamcap -uninstall

Have fun.