-------------------------------------------------------
CONTENTS: APPINFO, ABOUT, USAGE, TO DO, REVISON HISTORY
-------------------------------------------------------


[APPINFO]
name   : Procedure Speed Tester
version: 2003-11-04
author : redbird77
email  : redbird77@earthlink.net
www    : http://home.earthlink.net/~redbird77


[ABOUT]
This is a quickly built utility app that can run procedures a user-specified amount of times
and graphically (and via HTML) display the speed results.

The results are shown on the form like this:

Proc 1 - bar for normalized speed - time for all runs (time for individual run)
Proc 2 - bar for normalized speed - time for all runs (time for individual run)
...
Proc n - bar for normalized speed - time for all runs (time for individual run)

This app is not meant to super accurate down to the nanosecond.  It is just meant to give a rough idea on the speed of various procedures and more importantly the relative fastness or slowness of one procedure compared to another.

Although I recommend having the same number and type of parameters to each procedure in a group to get an accurate comparasion, there is no restriction on that.  Try the same procedure with ByRef instead of ByVal and see what happens to the speed (but be careful of evil side effects :).  Or try different data types or even Variants.

Also remember speed isn't everything, memory is an issue too.  A common technique to speed up code is using lookup tables for common values (like the powers of 2), but this can use a lot of memory.  Code wisely!


[USAGE]

(Note: see mGroup_Template.bas for a template.)

Remember to compile if you want your times to reflect running a compiled executable versus running in the VB IDE.

In mSpeedTester.bas
-------------------
In the sub SpeedTester_Run, change the first line to reflect where your procedures to test are located.

In mGroup_[InsertYourGroupNameHere.bas]
---------------------------------------
In the module add the code of all the procedures that you want to test.
In the RunTests Sub call all of those procedures.

In fSpeedTest.frm
------------------
Call SpeedTester_Init with your parameters.
Next call SpeedTester_Run and SpeedTester_Graph, and finally, if desired, call SpeedTester_ToHTML.


[TO DO]
Add the ability to indicate which procedures within a group to run (with checkboxes?).


[REVISON HISTORY]
2003-11-04
Fix - The oversight that would cause a "Subscript out of range" error.  (Thanks to Roger Gilchrist)
Add - The ability to automatically copy the HTML to the clipboard.  (Thanks to Robert Rayment)

2003-10-31
Initial release.