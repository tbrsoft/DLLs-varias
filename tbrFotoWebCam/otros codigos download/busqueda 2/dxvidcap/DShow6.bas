Attribute VB_Name = "mDirectShow6"
Option Explicit
'****************************************************************
'*  VB file:   DShow6.bas... for DirectShow 6.0 ActiveMovie control
'*  created 10/98 by Ray Mercer
'*
'*  This bas file includes functions for creating a filtergraph in VB5
'*  which will allow live preview from any VFW or WDM-compatible video
'*  capture hardware which supports preview.  Because of severe
'*  limitations in the quartz.dll's typelibrary, capturing to a file
'*  is not supported at this time.
'*  I did show how to connect to the Filewriter filter (check the
'*  commented code), but there does not seem to be any way to show
'*  a filter's property pages from VB!  This means we are allowed
'*  to connect to filters, but can't actually USE them :-(
'*  I am working on trying to figure out how to implement a TypeLibrary
'*  for VB which will allow access to more Interfaces, but for now
'*  if you want DirectShow VideoCapture you really should use C++
'*
'*  VisualBasic Functions for working with DirectShow 6.0
'*  DirectShow(tm) is a Microsoft API which can be DL from the
'*  Microsoft DirectX website at http://www.microsoft.com/directx/
'*  The DirectX Media runtime is required for this project to load
'*
'*  Copyright (c) 1998, Ray Mercer.  All rights reserved.
'****************************************************************
Public Function InsertFilter(ByVal FName As String, ByRef f As IFilterInfo, ByRef FGM As FilgraphManager) As Boolean
'returns false if the filter cannot be located

Dim i As Long
' call IRegFilterInfo::filter
    Dim LocalRegFilters As Object
    
    Set LocalRegFilters = FGM.RegFilterCollection
    Dim filter As IRegFilterInfo
    For i = 0 To (LocalRegFilters.Count - 1) Step 1
        LocalRegFilters.Item i, filter
        If filter.Name = FName Then
            filter.filter f
            Debug.Print f.Name
            InsertFilter = True 'return success
            Exit For
        End If
    Next i

End Function


Public Function CapPreviewConnect(ByVal DriverName As String, ByRef FilterGraph As FilgraphManager) As Boolean
Dim ret As Boolean
Dim capFilter As IFilterInfo    ' capture driver name
Dim PreviewPin As IPinInfo      ' preview pin
Dim RenderFilter As IFilterInfo ' video render filter for preview window
Dim RenderInputPin As IPinInfo  ' render input pin

'used in the remarked section of code which connects the capture pin
Dim CapturePin As IPinInfo
Dim AVIMux As IFilterInfo
Dim MuxInput As IPinInfo
Dim MuxOutput As IPinInfo
Dim FileWriter As IFilterInfo
Dim FileIn As IPinInfo

'used if you unremark the video effects code below
Dim EffectFilter As IFilterInfo
Dim EffectIn As IPinInfo
Dim EffectOut As IPinInfo

'First initialize the capture filter
ret = InsertFilter(DriverName, capFilter, FilterGraph)
If Not ret Then
    Exit Function
End If

'then find and connect the preview pin
Call capFilter.FindPin("Preview", PreviewPin)
If Not PreviewPin Is Nothing Then
    Debug.Print "obtained preview pin"
End If

''the following code works fine
''I can connect the Image Effects Filter to the filterGraph
''But there is absolutely no way to control it from VB!
'ret = InsertFilter("Image Effects", EffectFilter, FilterGraph)
'If ret Then
'    On Error Resume Next
'    For Each EffectIn In EffectFilter.Pins
'        If EffectIn.Name = "XForm In" Then
'            Debug.Print "obtained effect input pin"
'            Exit For
'        End If
'    Next
'    For Each EffectOut In EffectFilter.Pins
'        If EffectOut.Name = "XForm Out" Then
'            Debug.Print "obtained effect output pin"
'            Exit For
'        End If
'    Next
'End If

ret = InsertFilter("Video Renderer", RenderFilter, FilterGraph)
If ret Then
    On Error Resume Next
    For Each RenderInputPin In RenderFilter.Pins
        If RenderInputPin.Name = "Input" Then
            Debug.Print "obtained render input pin"
            Exit For
        End If
    Next
End If

''use this if you unremark the code above
'Call PreviewPin.Connect(EffectIn)
'Call EffectOut.Connect(RenderInputPin)

'otherwise use this
Call PreviewPin.Connect(RenderInputPin) 'connects the vidcap's preview pin to the VideoRenderer's Input pin

'the following code connects the pins correctly from the vidcap ~capture pin to the AVIMux to the Filewriter
'but there is no way to set the Filename for the FileWrite filter to write to!  So you get an automation error
'when you try to run the filter graph

''now that the preview pin is connected, find and connect the capture pin
'Call capFilter.FindPin("~Capture", CapturePin)
'If Not CapturePin Is Nothing Then
'    Debug.Print "obtained capture pin"
'End If
''insert the AVI multiplexer filter
'ret = InsertFilter("AVI Mux", AVIMux, FilterGraph)
'If ret Then
'    For Each MuxInput In AVIMux.Pins
'        If MuxInput.Name = "Input 01" Then
'            Debug.Print "obtained Avi Mux Input 01 pin"
'            Exit For
'        End If
'    Next
'    For Each MuxOutput In AVIMux.Pins
'        If MuxOutput.Name = "AVI Out" Then
'            Debug.Print "obtained Mux AVI Out pin"
'            Exit For
'        End If
'    Next
'End If
''insert the file-writer filter
'ret = InsertFilter("File writer", FileWriter, FilterGraph)
'If ret Then
'    For Each FileIn In FileWriter.Pins
'        If FileIn.Name = "in" Then
'            Debug.Print "obtained file writer input pin"
'            Exit For
'        End If
'    Next
'End If
''connect capture graph
'Call CapturePin.Connect(MuxInput)
'Call MuxOutput.Connect(FileIn)

CapPreviewConnect = True

End Function

Public Function EnumVideoCapHW(ByRef cb As ComboBox, ByRef lbl As Label) As Long
   'returns number of Video capture devices
   ' and loads their names into the combo box for user to select
   ' provides UI feedback through the label control
    
    '*Note we 'must create new FilterGraph because there is no way to enumerate certain types of filters without
    ' connecting them to a filterGraph, and there is no way to disconnect pins from the FIlter graph
    'in VB!
    Dim tempFIlterGraph As FilgraphManager
    Set tempFIlterGraph = New FilgraphManager

    Dim filterRef As IRegFilterInfo 'registered filters (only info we can get from this is the name as string!)
    Dim f As IFilterInfo 'gives a little more info about the filter, but does not provide any way to control the filter!
    Dim PreviewPin As IPinInfo
    Dim CapturePin As IPinInfo
    Dim numdevs As Long

'    Dim FilterCol As Object
'    Set FilterCol = FilterGraph.RegFilterCollection
    cb.Clear
    For Each filterRef In tempFIlterGraph.RegFilterCollection
        'loop through all filters' pins
        'if the filter has both a preview pin and a ~capture pin
        'then assume it is a vidcap filter and add it to combo-box
        'this is the only way to find this info from VB currently :-(
        Set f = Nothing
        lbl.Caption = filterRef.Name
        lbl.Refresh
        If filterRef.Name <> "Oscilloscope" Then 'kludge to avoid the stupid oscilloscope filter that appears as soon as you instantiate it!
            filterRef.filter f
            For Each CapturePin In f.Pins
                If CapturePin.Name = "~Capture" Then
                    'note* not all video capture devices will have a Preview pin
                    For Each PreviewPin In f.Pins
                        If PreviewPin.Name = "Preview" Then
                            numdevs = numdevs + 1
                            cb.AddItem f.Name
                        End If
                    Next
                End If
            Next
        End If
    Next
    lbl.Caption = ""
    If cb.ListCount > 0 Then
        cb.ListIndex = 0 'show first device
    End If
    EnumVideoCapHW = numdevs
End Function

