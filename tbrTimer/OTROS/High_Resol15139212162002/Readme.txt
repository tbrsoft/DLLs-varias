hTimer 1.0 Documentation

Contact Info.
hTimer was created by Adam Black.
For questions or comments you can email adz8@softhome.net
For freeware software visit my website: http://xarsoft.cjb.net

Table of contents
NOTE. You can choose to do either step 1 or step 2. Do not do both of these.

1. Compile hTimer.vbp into a dll and use the dll file in your project.
2. Import hTimer into your project instead of compiling and using the dll.
3. Using the timer in your project


1. Using the hTimer.dll file in your project
Note. All the following files you are told to locate should be found in the same
directory as this text file.

Open the hTimer.vbp project. Compile it to your system folder as "hTimer.dll"
Click Start, Run, and type. "Regsvr32 hTimer.dll". Press enter and your done.
In your project click the "Project" menu and then click "References".
Find the hTimer dll file in the list and add it to the references.

The advantages of option 1 is it makes the project easier to manage. You don't
have lots of extra Classes and modules hanging around.



2. Compiling the hTimer into your project
Note. All the following files you are told to locate should be found in the same
directory as this text file.

Click the "Project" menu, then "Add Class Module", click "Existing" and locate "hTimercls.cls"
Click the "Project" menu, then "Add Module", click "Existing" and location "hTimerBas.bas"
click the "Project" menu and then click "References", click the "Browse" button
and locate "win.tlb"

The advantages of option 2 is that if you compile the class and module into
your application, when you compile the app, no dll files are needed for the 
hTimer activex.



3. Using the hTimer in your project.
Now that you have used either Option 1 or Option 2 you are no ready to use the timer.
In the Declarations section of your form add the following statement

Dim WithEvents MyTimer As hTimerCls

where MyTimer is the name of the timer.

Now in the Form_Load Sub put the following code.

Set MyTimer = New hTimerCls

You are ready to use the timer
MyTimer.Interval = [Interval] 'Use this property to set the interval
MyTimer.Enabled = [True/False] 'Use this property to Enable/Disable the timer.





hTimer was created by Adam Black.
For questions or comments you can email adz8@softhome.net
For freeware software visit my website: http://xarsoft.cjb.net
