                 
   COMMUNICATIONS VERIFICATION AND TROUBLESHOOTING
   Win95

   Note: These instructions are also located in the Disk1, "Instrument
   Communication" folder of the install disks.


1. Make sure that the connector cord is plugged into the serial ports
   of both the PMI instrument and the computer.  Turn OFF both the
   instrument and the computer.  Turn them both back ON and reenter
   Windows 95.

2. Verify that HyperTerminal (Win95 component) is installed and
   working on your computer.

3. Open Win95 George.ht located in the CapWin directory, "Instrument
   Communication" folder.  Look in the floppy install Disk 1 if
   necessary.

4. Type "VAX".  If "VAS" (or "VAM") appears, communication with the
   instrument is successful.  Go to step 5.
   [Note that the third letter you type is "X" while the third letter
   on the screen is "S" (or "M").]

   If you get anything other than "VAS (or "VAM"), proceed as follows:

   Go to the File menu and select Properties.  Select "Connect using"
   in the Phone Number tab and change "Direct to comm 2" to a different
   comm port.  Close the box and try typing "VAX" again.  Keep trying
   different ports until you succeed. Go to step 5.

   If you have tried all of the comm ports and none work, call PMI or,
   if you have Microsoft Word, locate cap_bugs.doc in your Capwin 
   folder or install disks, print it, fill out the form, and fax it to
   PMI. Phone: 800-825-5764(USA/Canada), 607-257-5544.  Fax: 607-257-5639


5. Close Win95 George.ht.  Reenter CapWin.  Go to Calibrate, Select
   Comm Port, and choose the port that worked above in Win95 George.ht.
   Then go into Manual Control.  WARNING! YOUR COMPUTER MAY HANG.  Make
   sure you do not have any open programs with unsaved data. If it does hang
   reboot it. 

   If you cannot go into Manual Control, then you must try another
   comm port.  Go back to step 4.