Attribute VB_Name = "mTAPIvb"
'****************************************************************
'*  VB file:   TAPIvb.bas...
'*             VB Callback Proc for TAPI
'*
'*  created:        1999 by Ray Mercer
'*
'*  8/25/99 by Ray Mercer (added comments)
'*  3/09/2001
'*
'*  NOTE*  This callback proc REQUIRES that you pass an OBJPTR to
'*  the currect instance of the CvbTAPILine class when calling lineOpen()
'*  This is one method of simulating how C++ programmers can pass a Me pointer
'*  in the dwCallbackInstance.  However, in Visual Basic callbacks are volatile
'*  and if you try stepping through this callback you will probably crash your OS.
'*
'*  Copyright (c) 1999-2001 Ray Mercer.  All rights reserved.
'*  Latest version at http://www.shrinkwrapvb.com
'****************************************************************
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                                    (dest As Any, src As Any, ByVal length As Long)

Public Sub LineCallbackProc(ByVal hDevice As Long, _
                                ByVal dwMsg As Long, _
                                ByVal dwCallbackInstance As Long, _
                                ByVal dwParam1 As Long, _
                                ByVal dwParam2 As Long, _
                                ByVal dwParam3 As Long)
    'the callbackInstance parameter contains a pointer to the CvbTAPILine class
    'this sub just routes all callbacks back to the class for handling there
    Dim PassedObj As CvbTAPILine
    Dim objTemp As CvbTAPILine
    Debug.Print "LineCALLBACK : dwCallbackInst = " & dwCallbackInstance
    If dwCallbackInstance <> 0 Then
        'turn pointer into illegal, uncounted reference
        'Debug.Print "step #1"
        CopyMemory objTemp, dwCallbackInstance, 4
        'Assign to legal reference
        'Debug.Print "step #2"
        Set PassedObj = objTemp
        'Destroy the illegal reference
        'Debug.Print "step #3"
        CopyMemory objTemp, 0&, 4
        'use the interface to call back to the class
        'Debug.Print "step #4"
        PassedObj.LineProcHandler hDevice, dwMsg, dwParam1, dwParam2, dwParam3
        'Debug.Print "step #5"


    End If

End Sub



