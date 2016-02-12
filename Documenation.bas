Attribute VB_Name = "Documenation"
Option Explicit
'2-12-2016
'version Capwin 6.74.275.exe now calibrates pressures gauges on form load (calls CalibratePressures). This can also be done via the menu Calibrate > Calibrate Pressure Gauges.
' 6-10-2014
' Added the "LatchingValves" value to capwin.ini.  Value must be a comma seperated list of numbers or a single number
' any numbers listed will make those valves fire backwards (dumb latching valves) 0 turns off all of them.  Values are
' adjusted for 0 offset inside of the software
' 8-12-2013
' removed all references to VBCrypt.dll, and removed UAC features.  We ain't need this stuffs.
' 8-9-2013
' Removed path checking in Titlescrn.frm.  Will always default to the executable path.
' Added button to titlescrn.frm to "reset USB connection".  This is for new USB machines that work somtimes
' and not others.  This just resets the COM port.
' 9-26-12 version 6.74.101
'  Added initial testing support for Athena temperature controllers.  To use this, set up the temperature system
'  as if it were Modbus enabled Watlow(s), and then set Athena to 2 to signify that the watlow(s) have been replaced
'  by Athena 2 channel controllers.  This will have to be updated if we ever do 4 channel Athena controllers.
'  This has not been tested yet.
' 8-31-12 version 6.74.100
'  Added a funcion to set a minimum and maximum starting parameter to pressurehold
'  This new parameter will not appear on the log nor will the results of its test
'  Added a new error message for error number 8 instrucing the user on fixing the problem
'  Changed the color on some buttons and windows to a more appealing choice
'
'  TODO: minor code cleanup and stub removal,
'  Create an ending parameter,
'  link the parameters to the test's coded fail rate
'
' 7-30-12 version 6.74.99
'  EXPERIMENTAL - none of this has been tested on a real machine yet
'  Allow pressure hold testing in a bubble point tester.
'  Do the rotating chamber fill before the pressure hold test and drain after the test
'  Fixed: Pressure hold test would get an overflow error if the holding pressure was really small
'   and the minimum pressure in the regulator calibration table was rather high.  (It wanted to "increment"
'   the regulator to a negative value, which was so far negative it couldn't be used as an integer parameter)
'  TODO: multiple sequential pressure hold tests

' 06-19-12
'  07-12-12
'  Implemented bubble point (BP) finding new algorithm for cartridge tests.
'  It uses Ron's proposed pressur, a stack of measured pressures, instead of F/PT to determine
'  the pressure spread, and assigning the pressure, which corresponds to the
'  maximum pressure spread, to the BP pressure.
'  Implemented Methods:  InsertNewValue, getMeanValue, getStandardDeviation
'  Implemented global variables: m_bBPFindingForCartridge, BP_UsePressureVsTime,
'                 (If the first one activates Pressure Spread method, the second
'                 one change the progress screen contents and shows Pressure vs Time graph).
'                 User should set them "Y" in the capwin.ini file to activate them
'                 (defaults are "N").

' Run-time error 384 issue is fixed for the QC mode.

' 6.74.98
'  06-18-12
'    Two chamber bubble point tester software for the cartridge testing is working autotest mode.
'    Two air inlets and regulators test issue is fixed.

' 6.74.97
'   5-18-12
'     Added support for rotating samples using an extension of the auto wetting pump motor control
'     Note:  You can't have auto wetting and rotating chambers in the same machine

'6.74.96
'  Hydrohead test selection OptionButton is connected to the form,
'  which has three different test options: Run as Hydrohead test, Run as
'  Burst test, and Run as Mullen test. The Austin machine software is the
'  first implementation of this - 05-04-2012 (Surik Mehrabyan)

'Feb 4, 2011
'  Changed the progress screen so that the Turn menu option now says Next Step per customer requests.

'Feb 1, 2011
'  Added in new User Access Control features.  This allows the use of different user accounts that
'  can be logged in with.  Each user can be set as an admin account with full control or a user
'  account with limited control.  This does not replace the user of groups.  Groups are still used
'  to determine what test options are available.
'  Still need to add the ability to assign a specific group to a user if wanted.

'6.73.52
'  Altered the manual control screen to add in a third flow meter on the porometry side.  This is
'  for a new Seika machine, but will probably become standard at some point.  Altered the ReadXReturnX4
'  and raw_reading functions to utilitize this meter properly.  Added a new variable named xhflow_meters
'  that will determine how many extra flow meters are in the system.  Theoretically, this could be increase
'  later if you wanted to add a 4th in a similar way (not like Corning).  If this value is set to 1 and
'  the feature number determines that it is an xhflow machine, then there will be 2 flow meters.  If the
'  value is set to 2 then there will be 3 flow meters.  It tells how many extra meters there are.

'6.73.35
'  Fixed the auto wetting controls for all tests.  Removed the wetting_time variable.  It was
'  duplicating the auto_wet_wet_time variable.  Also, made changes to IF statements so it wet
'  at the right time for the right tests.
'  Have also added in new manual control screen for the corning machine called ManualControl1
'  and added all code to make that work.
'  Added code for 3rd MV for corning.

' 6.72.020 - rvw 7-23-10
'  ReserveTankLevelChannel is now used independently of the recirculation system
'  This allows a level sensor in the water tank for normal systems that have a water tank
'  Set this to the default of -1 to disable the display
'  Set this to 34 (low range pressure gauge 4 input) for new liquid level sensor
'  If you are using the pressure gauge 4 input for some other purpose, you will
'  have to choose a different channel for the liquid level sensor.

' 6.72.019 - rvw 7-7-10
'  Added resin test for Rice University.
'
' 6.72.017 - rvw 4-23-10
'  Added Resin_Diverter_Valve (defaults to 0 to disable this) for new Resin Intrusion system
'  for Rice University.  Only manual control for now.
'  If they have a Resin Diverter Valve, they also have control of a vacuum pump on the door lock
'  line (so you can't have a door lock and a Resin Intrusion system on the same machine).

' 6.72.016 - rvw 4-9-10
'  Added configuration values Num_Microflow_Volumes (defaults to 1 for normal operation)
'   Microflow_Volume2_Valve (defaults to valve 22 if Num_Microflow_Volumes > 1)
'   Microflow_Volume3_Valve (defaults to valve 23 if Num_Microflow_Volumes > 2)
'   Microflow volumes are still stored in array Diff_Volume.
'   You can't have Num_Microflow_Volumes > 1 if you are also using manualMultiChamber or
'    dry_chambers>1 (as this was the original use of Diff_Volume)
'
'       6.72.004 - JF 2-3-10
'                Added in code to run a Pass/Fail test based off of Median Pore Size.  This
'                funcionality requires opening CapRep right after the test finishes.  Half of
'                the code for this is in CapRep.

'       6.72.003 - JF 1-25-10
'                Changed code so that certain menu functions would not show up during BP tests.
'                Added in code to run multiple BP tests in a row.
'
'
'   Software Ver.
'       6.71.136 - AJB 12-23 to 12-24-09
'           .01) Added AJB_UTIL module to hold new code I add so it can be track down easier.
'       6.71.135 - AJB 12-15-09
'           .01) Added parameters to store third and forth flow meter.
'           .02) Added parameters to indicate that a system has a second motor valve.
'           .03) Added parameters open & close limits for second motor valve.
'           .04) Updated manual control to allow user to select which motor valve to open, stop, close, increment and decrement.
'           .05) Updated manual control to hide tank level controls properly.
'           .06) Updated manual control to display flow rates with commas to make it easier to read.
'           .07) Verify the low and high ranges for FM# 3&4 work, when user has selected MV2
'           .08) Close MV2 at beginning of test.
'           .09) Changed variables points, FPOINTS and SPOINTS to Long from Integer to prevent overflow error with long microflow tests.
'           .10) Fixed bug with the gurley option being available for liquid perm machine.
'           .11) Cut wetting code from Ron into the current code.
'           .12) New ini file variable "number_of_pistons", which defaults to 0
'           .13) If "number_of_pistons" is > 0, new ini file variables "piston_valve_1"
'               defaults to 15 (the normal piston valve).  2 defaults to 26 and
'               3 defaults to 27 (two new valves that were added just for this purpose)
'           .14) When operating a piston, the chamber number (for multi-chamber machines)
'               will determine which valve to use.  If there are more chambers than there
'               are pistons, the default piston will be used (15).
'           .15) New ini file variable "number_of_wetting_valves", which defaults to 0.
'               and "wetting_valve_1" through "wetting_valve_3", which default to
'               valves 22, 23, and 24.  This turns on a button on manual control which
'               brings up a new form for controlling these valves.
'           .16) New ini file variable "wetting_time", which defaults to 0.
'               If it is >0, and number_of_wetting_valves is >0, the wetting pump will be
'               activated during dry up / wet up test in middle instead of raising the
'               piston.
'           .17) Fixed problem with multi-chamber testing (where only first chamber would
'               be run over and over again) that was introduced when sequential testing
'               for special auto-wetting auto-draining system was added and then removed.
'           .18) New ini file variable "no_chamber_bypass" defaults to 0.  When set to 1 it
'               notes that you can't bypass the sample chamber when doing calibrations
'       6.71.134 - AJB 12-04-09
'           .01) Added module John Fish created to calculate a pressure from a count value and temperature for a 500 psi gauge.
'           .02) Added highTempPressureGauge variable to determine if the function is used.
'       6.71.133 -
'
'       6.71.132 - AJB 11-05-09
'           .01) Added tank fill form and logic to Load_Sample and Refill_Penetro functions. Tested and works.
'           .02) Added piston debugging function to manual control.
'       6.71.131 - AJB - 10-22-09 to 11-03-09
'           .01) Added logic for eureka machine to switch between main pressure gauge and gauge at the sample. LP test
'           .02) Added logic for eureka machine to switch between main pressure gauge and gauge at the sample. GP-type tests
'           .03) Added support for when temperature control is active to cool the chamber down by opening v2 and flowing gas.
'           .04) Added support to pause program execution until the chamber temperature has reached a desire set point. This
'                   is sample specific.
'           .05) Added support if a dock lock exists to prompt the user to open it, do what is needed to sample, close chamber
'                   and door and continue on with the test.
'           .06) Modified earlier gas test routine, works better now.
'           .07) Modified select_test form to display pressure step list form when the "Load Pressure Step list editor" button
'                   is pressed.
'           .08) Modified what was the load pressure step list button to now read, "Load Pressure Step List", from "c"
'           .09) Setup manual control screen to move sample pressure gauge labels into an ordered place and only
'                   display the correct number of pressure gauges labels.
'           .10) found a bug in the LP test where it opens v23 while the test is running and all the fluid
'                   runs out of v23 instead of going to the sample. Added ini option for topDownLP to determine flow pattern.
'           .11) Found bug in changes made to setup_2_3 where it was not going back to load sample for dry/wet test.
'           .12) Created TankFillDialog to popup, if tank_level_exists, to tell the user to press re-circulating pump switch.
'           .13) Added minFPT variable to capwin.ini file to allow a user to define a minimum F/PT value, lower than the
'                   default 50.
'
'       6.71.130 - AJB - 10-21-09
'           .01) Added logic to leak test ramp up to determine if the pressure is going up, if not alert the user and
'                   give them the option to ensure the gas is on or to quit.
'       6.71.129 - AJB -10-19-09
'           .01) Add support to log to watlow readings into CFT file. To enable define dryChamberTemperature and
'                   airTemperature as the channels desired. dryChamberTemperature and airTemperature are recorded in order.
'                   i.e. flow, pressure, drytemp, airtemp
'       6.71.128 - AJB - 10-09-09
'           .01) Fixed bug with manual control not displaying the true position of valve 11 when
'                   low pressure gauge is off scale.
'           .02) Made percent porosity invisible.
'           .03) Fixed manual control bug with V1 and switching regulators.
'
'       6.71.127 - AJB - 10-06-09
'           .01) Fixed syntax errors in run_lohm_cal so variables can actually affect the routine.
'
'       6.71.126 - AJB - 09-28-09
'           .01) Added BuildCakeByPressureAbort and fixed bug with hydrohead and diffpg.
'       6.71.125 - AJB - 9-28-09
'
'           .01) Made changes setting the flow controller for bub point tests.
'       6.71.124 - AJB - 09-24-09
'           .01) Made changes to manual control screen per Dr. G
'
'       6.71.123 04/16/09
'            New code for India-1752 works mechanically, but reporting for three new liquid perm choices
'            still needs reporting (or avoiding reporting in the case of the slurry tube wash cylce.
'           .01) Added "piston_position_transducer_exists" and "slurry_tube_exists".
'                Updated Get_Capstuff to load new system variables.
'                Updated load_user_global_stuff to loaded new user variables.
'           .02) Expanded Global command_issued% from 74 to 83 to handle the  new commands.
'           .03) Updated ManualControl ValveClik_Click to generate commands 75 through 80.
'           .04) Expanding PX1, PX2, PY1, PY2 from 7 to 11 for piston_position_transducer and slurry_tube_pressure.
'           .05) Updated Gage Reading in run_manual_control to handle piston position transducer and slurry tube level.
'           .06) Updated ReadXReturnX4 to handle "piston_position_transducer_exists" and "slurry_tube_exists".
'           .07) Expanded ts$ and added text for piston_position_transducer and slurry_tube_level
'           .08) Updated ManualControl Form Load to show piston_position_transducer, slurry_tube_level, and
'                slurry_tube_pressure when needed and show the new valves and motors on a new frame called "slurry_tube_frame".
'           .09) Updated run_manual_control output handling to drive the new devices.
'           .10) Added manual speed control for the Slurry Wash Pump including displaying pump flow calculation.
'                Added slurry_wash_pump_max_flow_cc to calculate flow based on 0-5V on AOUT.
'
'           .11) Added variables and calculation for Slurry Tube Level in cm and flow by change in level.
'           .12) Modified Select_Test form to have three new liquid perm choices show if slurry_tube_exists.
'                Modified logic in form to account for now having five choices; not just two.
'           .13) Fixed existing bug where selecting Amb would make the PListCheck invisible, but did not uncheck it.
'                Now, selecting any other liquid perm than Elevated Pressure will uncheck and make it invisible.
'           .14) Added new test parameters. Updated TestScrn to use them.
'           .15) Expanded qcshow and updated load_user_global_stuff for new versons of it.
'           .16) Updated save_user_unit_stuff to retain test parameters.
'           .17) Expanded Vpos() to handle new devices, but retreated from the highest numbers when it
'                looked like the rabbit did not want to talk with valves greater than 31 that got into lower case
'                letters in the ASCII chart.
'           .18) Updated RunTest to show new liquid perm choices.
'           .19) Handle closing V29 for slurry_tube_exists like V9 for H2OPERM.
'                This is because they share the same inlet from the air side and you want to make sure it is closed.
'           .20) Added code for three new liquid perm choices. Inserted two into RunTest just after the Load_Sample call.
'                Need to avoid Load_sample for new Slurry Tube Wash Cycle.
'           .21) Updated Load_Sample to only call for filling the penetrometer based on being liquid perm for amb or elev
'               (not the new choices).
'           .22) Widened status screen.
'           .23) Added set_slurry_wash_pump_aout_by_flow and zero_SlurryPumpSpeed
'           .24) Updated Load_Sample to only call for pre-filling the sample chamber based on being liquid perm for amb or elev
'               (not the new choices) where it was checking for not being autofill. The three new choices do not need to start liquid full.
'
'
'       6.71.122 1/28/08
'           1) If you run a dry up/wet up test with no sample load prompts, you will
'               now only be prompted once to enter the saturated sample - there will be
'               no second prompt after the zeroing of the flow meters.
'       6.71.121 12/28/07
'           1) Fix small bug to disable the 2nd chamber ready button if there is no 2nd chamber.
'       6.71.120 12/20/07
'           1) Sequential testing on two chambers should now be working.
'           2) File name should -Increase++ with each sequential test.
'       6.71.119 12/14/07
'           1) Wet, soak, drain times should now be working on autotesting, if enabled
'       6.71.118) 12/14/07
'           1) Added manual contol for 2nd chamber as defined in the capwin.ini
'       6.71.117) 11-29-07 -started
'           1) Multiple chamber(two at this point) contenues testing. Added the selchambers form
'               for the sequantial testing.
'           2) Add the Auto Wetting feature Manual Control only if the technical feature is enabled on
'               the ManualControl Form. Made the ManualControl Form Larger to fit the new feature.
'           3) Now supports integrity flow meter and auto compression on version 9 (problem with special
'               case for version 6 that is no longer a special case in version 9).
'           4) Dual air inlet now works properly for compression pressure regulator calibration
'       6.71.116)
'           1) Fixed OVERFLOW error caused by improper BARR measurement conversion back and forth.
'           2) Created the Auto Wet Feature tab in "presForm" that is currently in the testing state.
'       6.71.115) 11-19-07
'           1) Fix problem in liquid perm only machines if fy2(0,0)=0 and fy2(1,1)=0 where you can't
'               run elevated test unless you actually select it every time - it would default to
'               an illegal type of test and not prompt you for the elevated pressure values.
'           2) New parameter MinBubFlow - minimum bubble flow value for parameter editor.  Defaults
'               to 10 or 5 depending on what pressure gauges are installed.
'       6.71.114) 11-06-07
'            1) Chanced the bublflow values in the autoparm (Form)
'       6.71.113) 11-06-07
'           1) Now allow penetrometer fill to stop a little early - don't need to go all the
'               way to 2000 counts, can stop at 3200 (same tolerance we have on overfilling)
'       6.71.112) 10-22-07
'           1) Microflow volume calibration now works with autopiston machines
'       6.71.111) 10-12-07
'           1) Finished commenting out all of Ed's attempt at multi gas perm test so it will compile
'               again.  There are still some check boxes and other GUI things that are present, but
'               they are all invisible and any references to them are commented out.
'           2) Fixed subscript out of range error that would happen on systems with dry chamber
'               temperature but no reservoir temperature when you run any two-pass test (cfp).
'       6.71.110) 09-18-07
'           1) Added "pen_max_counts" to replace constant of 20000 or 62000 in determining what the
'               maximum readable count value is for a penetrometer
'           2) added code to change the bubble point flow minimum in the auto parameter form.
'       6.71.109) 09-13-07
'           1) Added support for two-color status light to show if test is running.  Set
'               "Status_Lights_Enable=Y" to turn this on.
'       6.71.108) 06-25-07
'           1) Fixed problem where if you enable multiple averaging for gas permeability this will cause
'               most porometry tests to end without saving data or zeroing pressure or closing valve 2.
'           2) Removed support for "special_ambient" test since this was only used in one porometer that had
'               two penetrometers where one of them was lowered, and this has since been replaced by the
'               ability to do a type of ambient test in a lowered penetrometer.  This was getting in the way
'               of the new support for dual penetrometer systems.
'           3) New variables to support dual penetrometer system:
'               Second_Penetrometer_V9 is the valve number of the venting valve for the second penetrometer.
'               Second_Penetrometer_V12 is the valve number of the isolation valve for the second penetrometer.
'               Second_Penetrometer_V13 is the valve number of the fill valve for the second penetrometer.
'               These all default to the original numbers (9, 12, 13) for initial testing.  To support the old
'               dual penetrometer system, the second V9 needs to be set to 19.  The original system didn't have
'               multiple valves for 12 and 13 since it wasn't autofill and one of the penetrometers was raised
'               and so didn't have a valve 12, and the other was lowered.
'               "MAXLQFLOW2" for maximum liquid flow for penetrometer 2.
'               "liquid_lohm2" for liquid lohm value for penetrometer 2.
'               Internal variables for above changed to arrays indexed off of penetrometer_select
'               Fixed valve 3 in max liquid permeability calibration for liquid perm only machines
'           4) Removed test to hide pore size calculations in some special cases because it was hiding it
'               when it shouldn't.  Now, it may be showing it when it shouldn't, but that is less bad.
'           5) Check if last type of test can actually be run on this machine, and switch to something that
'               can be run if it is not possible.
'           6) Added new test to hide pore size calculations during porometry test dry curve, per KG
'           7) Fixed rare false pressure gauge warning that some machines would get when they first power
'               up.
'       6.71.107) 06-15-07
'           1) fixed a problems with the edit parameter form for the liquid permiability
'               only test, Bought to light by the NIST job
'           2) the startf parameter was off in the edit parameters form added some error handeling so
'               if the parameter in the default .tpf file is higher or lower than the
'               program wants it is checked and rewwritten to the file.
'           3) Initializing routine for flow controller modified if it is a reversed
'               flow controller because they react differently from forward flow
'               controllers.  This should allow lower flow rates to work better, and
'               may speed up test initialization.
'           4) Fixed end of test sound - it was not shutting down sometimes
'           5)Fixed the problem with the loquid Perm file and the max pressure display and for
'               the lp only I hid labels 4 and 5
'           6) Calibrate form in manual control now doesn't lock up system if you close from the
'               close button on the window instead of clicking on the "exit" button
'       6.71.106)05-30-07
'           1)Minimum for the bubble flow in Auotparam is changed from 2 to 5 for low pressure machines.
'           2)problem with the Startf parameter in default.tpf file,if the min is less than the max the
'             max is rewritten in the help box but not in the text box or the file.
'       6.71.105)05-22-07 edc
'           1)changed the displayed minimum for the buble point on the autoparm form and the minimum that
'               the text box allows the user to enter
'           2)changed the minimum that the F/PT allows the user to enter and the displayed minimum.
'       6.71.104)04-30-07
'           1) altered the autoscale section of the hold pressure test so that the scale does not  keep
'               expanding with each passing data point.
'           2) changed the aut scale so that it will not scale into the negative pressure valuesthis can
'               be easily altered later if a such values are required.
'           3) changed the scale format so that the scale is in whole numbers.
'       6.71.103) 3-6-07
'           Added support for changing languages from within the preferences form.
'       6.71.102) 2-7-07
'           1) Dry curve now always has a data point at 0 flow at atmospheric pressure.  This is to fix a
'               problem on high flow porometers where the first real flow rate in the dry curve could be
'               at a pressure higher than the bubble point pressure, and this would cause the bubble point
'               data point to be dropped.
'           2) Added support for new valve 16 - dual air supply selector valve - if "Air_Inlets" = 2
'               (defaults to 1)
'       6.71.101) 1-17-07
'           1) Fixed bug in form load of regulator calibration form.  It would crash the program if you
'               tried to do a regulator calibration.  This was added after 6.71.97.
'           2) Liquid permeameters now don't show max flow rate or valve 2 control in the parameters
'               form.  These parameters are not needed, and can cause a lockup if you try to edit them.
'           3) Fixed problem with high flow bubble point on machines with a low flow controller
'       6.71.100) 1-2-07
'           1) New configuration parameter "Dry_Chambers", defaults to 1.  If you have a system with
'               more than one dry chamber but only one wet chamber, you use this in place of the older
'               multi-chamber system.  Still leave "CHAMBERS" to 2 to signify that you have a we chamber
'               and one dry chamber system.  This adds support for a new valve 5 that opens in place
'               of valve 4 to let the air into the second dry chamber.  Note that this new valve 5 moves
'               directly, as opposed to valve 4 which moves indirectly because it needs to support older
'               systems with a 3-way valve where the "closed" position let the air into the dry chamber
'               and the "open" position diverted the air to the liquid system.  To do this, the manual
'               multi chamber select has been extended to handle this case, and all calls that directly
'               control valve 4 (except for manual control) have been changed to call a more general
'               Dry_Chamber_Control, which will deal with valve 4 and maybe 5.  There is no direct control
'               of valve 5, except for in manual control.
'           2) Current pressure gets shown during leak test
'           3) Piston now comes down after fill in liquid perm tests with autopiston.  Before, it would
'               only do this for autopiston machines that werer only liquid permeability, not combination.
'           4) Autopiston with safetydown set to N will still prompt for door closed.  (It wasn't for the
'               lohm calibration.)
'           5) If the regulator was switched to the second regulator due to high flow rates during the first
'               pass of the test, the second pass will continue to use the second regulator and will properly
'               use the regulator calibration values for the second regulator so it sets the regulator starting
'               point properly.
'       6.71.99) 12-15-06
'           1)Added sound effects to the end of tests so that the operator knows that
'             the test has ended the sound loops until the
'           2)Added gas permeability logging feature
'               CalcGP now is a function that returns the average darcy value or an error code.
'               CalcGP now has corrected viscosity for Helium taken from CapRep
'       6.71.98) 12-4-06
'           1) Beginning of Ed's changes to add colored borders around all forms
'           2) Fixed pressure drop test setup - would cause invalid type error
'       6.71.97) 7-24-06 rvw
'           1) Fixed data averager when using gas permeability files.
'           2) Trapped out possible divide-by-zero error in bubble point pass/fail check routine.
'           3) Added penet_refill_delay, defaults to 0, to wait for some slow ball valves to finish
'               moving after refilling the penetrometer.
'           4) Added support for three more temperature probes (for a total of 6).
'           5) Modified single pressure point test so it works better in very low pressure machines.
'           6) Modified end of bubble point routine when you are using microflow porometry test
'               so it won't save too high a flow rate for the bubble point.  This may also fix the
'               microflow bubble point routine and make it less sensitive to false bubble points.
'       6.71.96) 7-10-06 rvw
'           1) Now doesn't save tortuosity factor for file types that don't need it.
'       6.71.95) 5-29-06 rvw
'           1) Network connection has better error trapping.
'       6.71.94) 5-4-06 rvw
'           1) Network connection system added.  Need to have network_connection_enabled=1 in
'               configuration file.
'           2) BubblerLevelChannel now defined for reading bubbler level sensor in manual control.
'               Also needs BubblerLevelZero and BubblerLevelSpan.
'       6.71.93) 4-3-06 rvw
'           1) QC Mode now matched group changes with title screen
'           2) microflow_extraChamber is now used during microflow calibration
'           3) liquid perm debug log header now includes units (psi, cm, sec) to make it easier to
'               understand.
'           4) When using frazier chamber and frazier pressure gauge, do not use any lohm correction
'       6.71.92) 3-8-06 rvw
'           1) Added lperm_autoFillVentTime, which defaults to 5.
'           2) Added microflow_extraChamber, which defaults to 0.  When it is 1, this means that
'               there is an extra chamber for microflow and valve 4 needs to be switched as if
'               we are doing liquid permeability during microflow tests.
'           3) Fixed problem with hydrohead tests on temperature controlled machines.  Hydrohead
'               data files are not supposed to contain temperature information, and the current
'               report program does not know how to deal with temperature information in a
'               hydrohead data file.  We now no longer put out the temperature sensor header
'               on such files.
'           4) Edit parameter file button now disabled during liquid permeability tests (unless
'               it is a special ambient test)
'           5) Debug log for liquid permeability now stored intermediate readings of pressure
'               height and time.
'           6) New variable FrazierChamberValve (defaults to 0 to disable, set to valve number that
'               needs to open when you want to use the frazier chamber).  Otherwise, works the same
'               as the existing (but rarely used) "FrazierPressureGauge="Y"".  Autopiston can now
'               equal "F" to signify that only the frazier sample chamber has a piston.
'       6.71.91) 2-7-06 rvw
'           1) During bubble point test on dual regulator machines, if regulator switchover
'               causes a pressure drop, this won't trigger a false bubble point reading.
'       6.71.90) 1-23-06 rvw
'           1) During the elevated pressure liquid permeability setup, added some traps to the
'               regulator initialize routine so it doesn't lock up in stage 2 if the regulator
'               does not get back to low pressure when the regulator gets to 0 counts.
'       6.71.89) 1-3-06 rvw
'           1) Curve fit is now enabled for liquid permeability only machines.
'           2) Added more information to debug output for liquid permeability.
'       6.71.88) 12-20-05
'           1) Fixed gas and liquid selection and creation of data files with non-standard
'           liquid or gas.
'       6.71.87) 11-29-05ecjc
'           1)Made a modification for Pella that will allow the supervisors to dictate if a user
'           can alter a specified parameter in the elevated liquid permiability test. This is done
'           by adding a check box to the got value form which is visible only in the supervisors mode
'           and when checked the user will be prompted to provide this information befor the test is
'           run.
'           2) Liquid permeability maximum pressure parameter now saved properly. (rvw 12-9-05)
'           3) New parameter for elevated pressure liquid permeability test -
'               lperm_initializeRegulatorPressure, defaults to 0.5 PSI.
'           4) New parameter for elevated pressure liquid permeability test -
'               lperm_regulatorIncrementSteps, defaults to 1.
'           5) Elevated pressure liquid permeability no longer sets the temperature to zero at
'               the beginning of the test.
'       6.71.86) 11-18-05 rvw
'          1) Fixed problem where some incorrect flow readings could be put into the data file
'           right after a pressure regulator switch on a low flow sample if the
'           reg2_high_flow_switch_count parameter was too low so the program needed to increment
'           the regulator more to get back to the pressure of the last data point.
'       6.71.85) 10-31-05 ecjc
'          1) Modified the select Gas and Liquid form so that the program could call a single
'           form for the purpose of creating and modifing liquids and gases. This form gives the user
'           the ability to pick a gass or liquid from a drop down box and alter the surface tenseion
'           or the viscosity or add a new gas or liquid. these are stored in an external textfile
'           in the executable path.
'       6.71.84) 10-26-05 rvw
'        1) Valve 3 is now closed during liquid permeability tests when valve 23 is present.
'           Also valve 3 doesn't open when you resume from hold during liquid permeability tests
'           for liquid perm only machines.
'           Also fixed autofill routine so really slow fills will properly wait longer if the
'           penetrometer is still going up.
'       6.71.83) 9-15-05 rvw final 9-27-05
'        1) For I/P regulator machines (all current machines) regulator now has the option of
'           starting at the SHFP instead of SBPP before opening valve 2 at the start of the dry
'           curve or after the bubble point.  This helps with very low flow samples.  This may
'           make high flow samples take more data points at really low pressures and flows.  The
'           new configuration parameter "first_flow_starting_point_percent" defaults to 100,
'           which uses the SBPP and will result in no change in the way the test starts.  To start
'           the regulator at a lower point (to take more low pressure data) lower this parameter
'           to 10, which uses the SHFP.  This only works on IP Regulator systems (version 7 or 8).
'           SBPP is the regulator position that allows the flow controller to reach 100% of full
'           scale.  SHFP is the regulator position that allows the flow controller to reach 10%
'           of full scale.  You could set a value in between 10 and 100.  You may also be able to
'           set a value lower than 10 or greater than 100.
'        2) For version 7 and higher, if the feature number designates an integrity flow meter,
'           you can have the integrity flow meter less than the low flow meter.  (The integrity
'           flow meter range is used in place of the low range of the low flow meter, and this
'           used to be the only determination that there was an integrity flow meter.)
'       6.71.82) 8-25-05 rvw modifications in Japan after delivery of 81 to customer
'        1) Single pressure gas permeability now saves temperature data (for porometers
'           with temperature capability).
'       6.71.81) 8-18-05 rvw final 8-24-05 in Japan for delivery to customer
'        1) liquid_lohm is now used properly to correct for pressure drop based on
'           liquid flow rate during liquid permeability tests
'        2) Improved pressure control for Ambient liquid test for lowered penetrometers
'        3) Ambient liquid test now correctly records temperature for recirulation machines
'        4) Improved pressure control for Elevated liquid test
'        5) Improved control of valve 12 (liquid chamber isolation) so it doesn't open
'           when it shouldn't on recirculation machines to keep the chamber isolated while
'           the penetrometer is filling.
'        6) Fixed curve fit problem with data files with two temperature values
'        7) Name of pressure list is now stored in the extended header of data files (for
'           those that are run using the pressure list, which are only gas or liquid permeability)
'       6.71.80) 7-21-05 rvw, final 8-18-05
'        1) Regulator calibration on liquid perm machines now closes valve 4 and doesn't ask you to seal
'           the sample chamber.
'        2) Max Liquid Flow test now saves better data into text file (for debugging purposes) and also
'           works better with dual regulator systems.
'        3) You can now go from partial recirculate to full test running, bypassing the full recirculation
'           mode.  This is for samples that are already pre-loaded and where you don't want any flow through
'           the sample until the actual test starts.
'        4) Max Liquid Flow test now calculates a lohm value for the liquid flow.  This is stored in the
'           capwin.ini file as "liquid_lohm", which defaults to 0 to signify no resistance to liquid flow.
'        5) F/PT value for no change (or drop) in pressure changed from 99999 to 9999999 to make it work
'           better with high flow machine that can have actual F/PT values of 100000
'       6.71.79) 7-19-05 rvw
'        1) Lohm calibration on dual regulator systems now uses parameter "reg2_high_flow_switch_count"
'           when it switches over.  Improved switchover mechanism.
'       6.71.78) 7-11-05 rvw, final 7-14-05
'        1) Maximum liquid flow test now runs at slightly above 1 PSI differential air pressure instead
'           of ambient pressure.  This allows systems with lowered penetrometers to run this calibration.
'           During the calibration test, MAXLQFLOW.TXT is created to store intermediate values for use in
'           debugging the new lohm calculation for liquid flow.
'        2) Multi-chamber systems now have support for manual multi chamber, which means that the chamber
'           selection is done by the user manually switching the supply hose.  The only purpose of this
'           so far is to have multiple chambers with microflow, where the chamber isolation valves are now
'           used as inlet selector valves for the microflow pressure gauge.  The microflow calibration needs
'           to know which chamber is being used for the calibration.  Chamber volume and microflow volume
'           are now stored in an array indexed off of the chamber that is being manually selected.  The
'           chamber selector now knows that only one chamber can be selected for this type of machine.
'        3) New liquid drain valve (23) now works properly for recirculation machines.
'        4) Bubbler valve now opens after bubble point and before start of wet or dry curve.  It will work
'           for gas permeability test too.  It should stay open at end of "wet up/dry down" test, but this
'           is not tested as this type of test is not normally recommended.
'       6.71.77) 7-6-05 rvw, final version 7-7-05
'        1) Added support for cyclic compression on systems with door switches.  Previously, this would
'           only work if there was no safety feature or on autopiston systems which always have door
'           switches but which do not have automated regulators.
'        2) Added support for setting any valve command letter for the multi chamber isolation valves.
'           This defaults to a string of "abcdefghij", which stands for normal isolation valves 1 through 10.
'           For special cases where these valve mappings are already used by other special valves, such
'           as a multi-chamber microflow system, this can be changed in the configuration file.
'        3) Use of the bubbler is now optional.  If the bubbler is enabled in the configuration file,
'           you now have to turn it on in the preferences form or it won't be used during the test.
'       6.71.76) 6-17-05 rvw, final version 7-6-05
'        1) New valve 23 is now supported in liquid permeability tests.  It will open to drain out
'           excess liquid from the penetrometer at the beginning of the test instead of pressurizing
'           the penetrometer to force the excess fluid through the sample.
'        2) New valve 25 replaces dual use of valve 4 for bubbler.  Disables valve 22 (which was only
'           used on Ballard).  New configuration parameter "v22_exists", which defaults to "N".
'        3) Better support for dual regulator machines where the low pressure regulator can't get to
'           full flow rate.  New configuration parameter "reg2_high_flow_switch_count", which defaults
'           to 300.  This is the regulator count value that is used as a starting point when we switch
'           to the second regulator due to not being able to reach a high flow rate on regulator 1.
'        4) During lohm calibration, if the flow rate ever goes off scale, it won't record a lohm value
'           for that flow rate.  It had been doing this, and this could cause an incorrect lohm value at
'           the end of the lohm table.
'        5) After switchover to second regulator during lohm calibration, the rate of increase of the
'           regulator is slowed while hunting for the starting point so it doesn't overshoot as much due to
'           a slowly reacting system.
'       6.71.75) 6-14-05 rvw, ecjc, MJLC final version 6-17-05
'        1) Fixed integrity test so that if the flow meter ever gets switched from integrity to
'           something else (such as when you go to manual control and click on the low flow meter)
'           it will get switched back.
'        2) Added several lines of code to the CAPFLOW.BAS and the CAPMAIN.frm that enable the program
'           to pull up the new help system in the usersdefault browser. The material in the help system
'           has also been updated to reflect the newest version of the users manual. - ecjc
'        3) New optional bubble point time log (also works with hydrohead) creates a log file with
'           pressure, flow, and time values that can be imported into a spreadsheet.
'        4) MJLC: Updated lanugage files
'       6.71.74) 5-16-05 rvw, MJLC - final on 6-1-05
'        1) Added support for dual stage compression (uses new valve 24, "COMPRESSION=D")
'        2) This also adds new variables pretreat_flow and pretreat_time, both defaulting to 0 to disable
'           them.  Flow is in cc/min, time is in seconds.
'        3) Cleaned up lohm calculations.  It will now properly turn a test when the lohm ratio gets too
'           low again. (It wasn't doing this since we changed lohm functions and changed the meanings of
'           some of the internal variables.)
'        4) MJLC: Fixed a problem encountered when switching from a test using a pressure step list to one without:
'          points for the list would be interpolated at the beginning of the regular data file, creating a corrupted
'          report. Now the step list option is visible whenever it can be applied to a test, so the user should see
'          whether it is selected or not. Also added a check at the end of the test selection form to turn off the
'          step list option if it cannot be used with the test type chosen.
'           Also re-enabled the "show test results in capwin" option in the Preferences window -- and disabled the
'          temporary hack that always turned it on.
'       6.71.73) 5-12-05 rvw
'        1) Some key values are now required to be above 0.  If you enter a 0 or negative value, the old
'           value will remain.  If the old value is <=0, it will be reset to the default value.  These
'           parameters are surface tension, diameter, thickness, density, and mass.  Cyl_len is allowed to
'           be 0, as this is normal for flat sheets.  SCDiam is allowed to have any value because if it is
'           less than the normal diameter it is ignored.
'        2) Fixed problem in lohm calibration that would sometimes cause it to stop before it even started.
'       6.71.72) 4-25-05 rvw
'        1) Fixed elevated pressure test with one single very high pressure - if you got too close to the
'           maximum pressure it could try to set the gauge index to -1, which is not valid.
'       6.71.71) 4-21-05 rvw
'        1) Fixed curve fit so it now works with permeability files (this was broken in 6.71.70 when
'           support was added for temperature data files).
'       6.71.70) 4-18-05 rvw
'        1) Rewrote curve fit to use standard data file i/o routines.  This allows temperature data to
'           be used in the curve fit routine.
'       6.71.69) 4-12-05 rvw
'        1) (rvw) Fixed increment of I/P regulator during integrity test.  It would increment too much if you
'           tried to use a pressure regulator increment less than 1.
'       6.71.68) 3-16-05 rvw, MJLC final on 4-7-05
'        1) If you try to take more than 32000 data points, the test will automatically turn or abort.  This
'           stops an overflow that would happen at 32767 data points due to integer indexing of arrays.
'        2) If the open limit from valve 2 reads more than 65000 counts, it will give an error message.  This
'           has happened once when a valve pot board was shorted out.
'        3) Fixed some temperature reading problems with new temperature system and recirculation and liquid
'           permeability, and some door switch issues where old Ballard door switch code would interfere with newer
'           Rabbit door switch method (where it didn't interfere with older George door switch method).
'        4) Fixed some piston actuation problems in two-pass tests on recirculation machines
'        5) Fixed temperature sensor readings in porometry data file.
'        6) (MJLC) Fixed problem in display of pressure step list checkbox in test selection form. Was not being
'           displayed properly because a check of the liquid perm state was being called during the form load method,
'           and the default liquid perm state turned off the pressure step list option.
'        7) (rvw) Added button to data editor to strip temperature data from a file.
'        8) (rvw) Fixed invalid parameter error when loading auto parm form - caused by slider initialization and
'          bubble flow value > 30.
'        9) (MJLC) Added running statistics (avg and stdev) to data reported in log file for a porometry or BP test.
'       10) (MJLC) Fixed(?) a bug where the "dry parameter" label in test setup was indented relative to the others.
'       11) (MJLC) Added regulator switching to the pressure step list routine.
'       12) (rvw)  Added check for pressure and flow stagnation when using first of two regulators.  This allows
'           switching to the second regulator even before the first regulator reaches 100% if the flow rate is
'           so high that the first regulator can't increase the pressure or flow any more.  This is only working
'           for the lohm calibration for now, and could use some improvement.
'       6.71.67) 2-17-05 rvw
'        1) Added support for bubbler attachment.  This is turned on by "bubbler_enable=Y" in capwin.ini file.
'          Moves valve 3 to inside of valve 4, valve 4 leads to the bubbler, a second valve 3 is on the other
'          side of the bubbler.  Valve 4 stays closed except when actually running a test (to avoid back flow
'          of water from bubbler).  Valve 4 is still operated in reverse (close commands causes it to open) due
'          to historic nature of valve having been 3-way valve in earlier machines.  Valve 4 is now closed
'          at end of test when there is still pressure on the system.  After 3 seconds (to allow valve to actually
'          close) the pressure is then relieved and both valves 3 open to vent the system.  It stays in this
'          configuration until the next test is started.
'        2) Changed the way temperature probes are allocated.  Previously, the "Temperature" configuration
'          value was set to "N" for no temperature probes, "Y" for 1 temperature probe, or "2" for two temperature
'          probes.  The allocation of these probes depended on if the system had liquid permeability or not.
'          There was no support for the new Rabbit pass-through serial ports to talk to a Watlow.  Now, there are
'          three new configuration variables that replace the "Temperature" and "external_watlow_com_number"
'          configuration values.  "DryChamberTemperature", "WetChamberTemperature", and "ReservoirTemperature".
'          These default at "0", to mean that there is no probe in this section.  When set to "G1" or "G2", this
'          means the probe uses the old V6 and V7 aux port 1 or 2 (channels 27 or 30).  When set to "C1A", this
'          means that the probe is on channel A of a Watlow connected to the computer's COMM port 1.  When set to
'          "R1A", this means that the probe is on channel A of a Watlow connected to the Rabbit V8 board
'          pass-through COMM port 1.  If none of these are set in the configuration file, the old "Temperature"
'          and "external_watlow_com_number" values are read in and used to set an initial condition of these new
'          variables.  If these is a conflict when trying to set these values, a warning message will be shown
'          (such as when "Temperature" says there are two probes but the system doesn't have a liquid chamber).
'        3) New configuration variable "valve_limit_offset", which defaults to 0.  Set this higher if you are
'          having problems where valve 2 is not able to reach the close limit but always stays slightly above it.
'       6.71.66) 2-1-05 to 2-7-05, MJLC, RVW
'        1) Made the pressure increment for the initial increase in the microflow test a variable (xregstep in
'           capwin.ini). Previously hard-coded to be 10. This could probably be added to the Preferences form, but
'           hasn't been yet.        -- MJLC
'        2) Fixed multi-chamber select form - due to a bug in the way the language file was read in, all the
'           chamber checkboxes would be called "chamber 1".
'        3) Fixed multi-chamber support for version 8 machines.  In version 8, chamber select valves 1 through 5
'           use the same solenoid valve circuits as valves 4, 20, 19, 21, and 22.  These numbered valves are not
'           present in a multi-chamber system, but the valve commands can incorrectly move the chamber select
'           valves.  This is now blocked in the Move_Valve routine.
'        4) Added support for 2 chamber multi-chamber system.  Previously, chambers=2 signified that there was
'           a liquid chamber and an air chamber, and chambers>=3 signified a true multi-chamber system.  Now,
'           if chambers=2 and the machine does not do liquid permeability (or only does liquid permeability
'           and does not do gas permeability) then it will be considered a multi chamber system.  Not all
'           combinations of this have been tested.
'       6.71.65) 1-31-05 MJLC
'        1) Bug fixes for the free-pressure test (Saint Gobain's thing).
'JF CHECK HERE
'       6.71.64) 1-21-05 MJLC
'        1) Added support for a new type of test: a multi-test averaging gas perm routine. Preferences are set
'          in the Gas Perm tab of the settings pane. The methodology is similar to the loopingDemo functionality
'          that exists. User enters the number of tests he wants to run. The software automatically loops through
'          Run_C_Pass and RunTest without pausing between tests; each data file has a number appended to its name
'          and is saved as normal. A summary file with the extension "_summary.cft" uses CalcGP to figure out the
'          average darcy value for each result and print it to the file. At the end of the last test, the average of
'          the averages is calculated and written to the summary file.
'        2) Modified calcGP. This calculated an average darcy number as a weighted average with respect to flow. Caprep,
'          however, does it with respect to pressure. Dr. Jena and Dr. Gupta agreed that it would be better to standardize
'          around pressure. (Now we can start reporting the result at the end of the test again!)
'       6.71.63) 1-19-05 MJLC
'        1) Bug fix in pressure step routine: when the interpolated data file was created, the wet and dry
'          curves were saved with the same number of points as in the original file. If there were more points
'          in the pressure list than in the data file, the last interpolated values would not be saved.
'       6.71.62) 1-12-05
'        1) Automatic curve fit won't work on bubble point tests.  -- MJLC
'        2) Fixes made to pore diameter option (pressure conversion removed, trap for 0 diameter) by Ron.
'        3) Bug fix in pressure hold routine: crashed if more than 100 data points were recorded.
'        4) Statistics.bas (in common library) was updated to use long values instead of integer (again for
'          pressure hold test).
'       6.71.61) 12-14-04 MJLC
'        1) Added a wait time to the piston/compression close routine. This is meant to give the piston
'          time to fully close in systems where it may take a while. Set as piston_delay_time in capwin.ini;
'          variable is pistonDelayTime.
'        2) Fixed a bug in the pressure step list editor where lists with 10 or more points were not loaded correctly.
'        3) Wait for flow meter stability is now bypassed for liquid perm tests as well as pressure hold.
'        4) Added support for use of a pressure step list in elevated liquid perm tests.
'        5) Expanded the step list interpolation done at the end of a test for gas perm to include liquid perm and
'          cfp tests as well.
'        6) Tortuosity factor not printed in data file for permeability or pressure hold tests.
'       6.71.60) 12-10-04 MJLC
'        1) Changed the "test by pore diameter" option so that min and max values are displayed
'          correctly in the setup form when flipping back and forth between the two options.
'        2) Test setup form is now "dirty" if min or max pressure/diameter is changed.
'        3) Added an "experimental" label to the main form that becomes visible if "x" is appended
'          to the global version number (S_Version$). This should prevent accidental release of
'          untested code.
'        4) Added a new prefs tab, "Special Options". First special option is zeroTempAtEndOfTest,
'          which determines whether or not the temperature should be zeroed at the end of run_c_pass.
'       6.71.59) 12-10-04 MJLC
'        1) Bug fix in the Saint Gobain routine (x5-->x4)
'       6.71.58) 12-6-04 rvw
'        1) Added support for looping demo when running dry up/wet up test.  First loop will run
'            normally, but subsequent loops will re-run the dry up part of the test and then use
'            the data from the wet up part of the first loop.
'       6.71.57) 12-1-04 MJLC
'        1) Log file option preference was not being saved properly. Fixed.
'        2) Modified testscrn form_load method by setting dirty=false. This will prevent program
'            from asking user if he wants to save changes if all he's done is opened the form and closed
'            it again without touching any of the controls.
'        3) Disabled the pressure unit check at the beginning of save_user_global_stuff. The called routine
'            was moved to the preferences form some time ago, meaning that every time the save routine was
'            called it was reloading the preferences form in the background just to run the pressure check.
'            The check isn't needed here anyway since the units can only be changed now in prefs form.
'        4) During main flow test (either wet or dry), stop endless retrying to increase pressure if we are
'            already at 4000 counts on an i/p regulator.  This will stop tests when they can't go on any
'            further, and allow dual regulator systems a chance to switch to the higher regulator where they
'            would otherwise hang. (12-2-04 rvw)
'        5) Multiplier for pressure regulator increment is normally 50 (to compensate for differences between
'            version 6 and 7 machines).  This is now lowered to 7 for second regulator in dual regulator
'            machines to compensate for the second regulator being 7 times the maximum pressure.  Also added
'            delay during the change in regulators to let it settle better.  This may help stop a large
'            jump in the data when we switch regulators.
'       6.71.56) 11-30-04 MJLC
'        1) Disabled part of the new liquid balance routines because they were causing an overflow error
'            during a regular elevated liquid perm test.
'        2) Moved the ambient liquid perm test routine out of run_c_pass into its own routine (Run_Ambient_LqPerm)
'            in an effort to simplify and modularize run_c_pass.
'        3) Made some modifications for the special Saint Gobain test again.
'       6.71.55) 9-16-04 MJLC
'        1) More changes to the Saint Gobain test: 5 second delay after acquiring a point, changes to
'            the formatting of the excel spreadsheet.
'       6.71.54) 9-7-04 to 9-16-04 Ron
'        1) Fixed some things in elevated pressure liquid permeability that were messed up when the
'            new optional balance code was inserted for microflow liquid permeability.
'       6.71.53) 9-3-04  MJLC
'        1) Modified Run_C_Pass and UpdateLine25 so that pore diameter is only displayed if the test type
'            is porometry or bubble point.
'        2) Modified the special Saint Gobain "free pressure test" -- now includes more test parameters,
'            Excel file output, two methods of triggering a read, and permeability calculation. Test is still
'            only accessible by enabling the "Saint Gobain" button in Testscrn at compile time.
'       6.71.52) 8-30-04 MJLC
'        Fixed display problems in the single-point gas perm test, updated language files for same, and
'         made the "setup" button for it in the "select test" form functional.
'       6.71.51) 8-30-04 Ron
'        Leak test now sets the proper regulator on a dual regulator machine
'        It does this based on the maximum pressure for the leak test, at the
'        beginning of the leak test.  It does not try to switch regulators in
'        the middle of the leak test.
'       6.71.50) 8/26/04 TAR---Test_Done() open the drain code moved for liq perm
'       6.71.49) 8-25-04 Ron
'        Fixed problem added in 6.71.47-8 when running dry curve first could cause overflow on some
'         machines due to no bubble point being defined at the beginning of the dry curve.
'       6.71.48) 8-24-04 Ron
'        Fixed problem where bottom-up liquid permeametry with very low flow sample and compression
'         could cause pressure buildup during initial attempt at moving the penetrometer magnet into
'         range so that the compression piston would be lifted up by the pressure on the sample.
'       6.71.47) 8-19-04 Ron
'        1) Fix bug in update_units_check.  If they had previously selected a non-standard unit for
'            pressure and that unit is not defined any more (maybe they changed the capwin.ini file
'            because they switched machines) they would get an error when they tried to load the
'            preferences form because it couldn't find the unit it thought they wanted.  Now it just
'            defaults back to PSIA if it can't find the unit they wanted.
'        2) Added language support to the two new items added in 46
'        3) Added support for compression piston in leak test, lohm calibration, regulator calibration,
'            and chamber volume calibration (piston was already supported in microflow volume calibration
'            so chamber volume calibration is now using the same routine as the microflow volume calibration
'            except it exits after the first pass).
'        4) Fixed problem seen in Brazil where conversion of tortuosity factor of .715 was causing
'            a "type mismatch" error.  All reading in of configuration settings that should be numbers
'            is now put through the "val" function to convert to a number using US standard.  (All
'            configuration settings are stored as US standard so configuration files can be moved
'            anywhere in the world.  Numbers are displayed and entered from the user using whatever
'            the local standard actually is.)
'        5) Fixed some places on the preferences form where str$ was used and it should have been
'            format$ to allow local decimal points.  Also, when these text values are read back in
'            they have to use myVal instead of Val so it will properly convert.
'        6) Fixed other areas on older forms where they used str$ instead of format$.
'        7) Fixed keypress restrictions on data editor to allow local decimal character
'        8) Fixed pressure regulator setting after bubble point that may have been setting the pressure
'           regulator too high for the first part of the wet curve
'       6.71.46) 8-18-04 Ron
'        1) New variable "min_flow_in_dry", defaults to 0.  If set non-zero, then dry curve going up
'            will skip data points in same way as "use_min_pressure_in_dry"
'           Also has "use_min_flow_in_dry" to turn it on
'        2) Also fixed bug introduced somewhere between 35 and 38 where liquid permeability would mess
'            up calibration values for low pressure gauge, thus messing up dry curves.
'       6.71.45) 8-12-04 MJLC
'        1) Added a new "single point gas perm test" for Sam Bo. There is an option for it in the
'            test selection window, as well as a Gas Perm preference pane in the preferences window.
'            The test uses the pressure step list pressurization routine to reach a target pressure
'            point, then holds that point for an assigned period of time, taking data points at specified
'            intervals. Currently it creates a test data file along with the normal file. **** THIS HAS NOT
'            BEEN COMPLETELY TESTED YET ****
'       6.71.44) 8/9/04 TAR---g_bBalanceNotPenet_zeroPoint added
'       6.71.43) 8/6/04 TAR---Added key-cancel to drain feature, moved drain feature in CAPFLOW.BAS.
'           Fixed Settling feature so that it doesn't eat into the sampling time.
'       6.71.42) 8/4/04 TAR---Added settling time for each press target, plus indicators "Waiting" and
'           "Settling". Now a new dialog parameter box comes up prior to Autotest for the settling time.
'           Added drain at the end of autotest. Fixed targeting at 0 PSI. Localized 10mL max for Mettler.
'           Debugging file output during Run_Elev_LqPerm. TAR_Util.BAS for Progress_Output line items. ptarg
'           double-checks. All sorts of good stuff.
'       6.71.41) 7-26-04 MJLC
'        1) Fixed bug introduced in previous version merge where hydrohead test option was never visible.
'       6.71.40) 7-23-04 MJLC
'        1) Updated language files for Mr. Yaza.
'       6.71.39)EdC 7-22-04 The slider(1) of the autoparm form was altered to effect only the
'        Bubblepoint aspect of the test.
'        Language files updated for new sliders and v.38x (MJLC)
'       6.71.38x) 7-22-04 Ron
'        1) Added support for exclusive hydrohead tester (setup for bubble point tester
'           and then set "hydrohead=E" in capwin.ini file).
'       6.71.38) 7-21-04 MJLC
'        1) Added an option to automatically curve-fit the data file at the end of a test. Done by
'            calling the curve fitting routine in Curve from do_final_copy. Preference added to the
'            "tests" tab of the preferences form. Also added a "curve fit" tab to preferences with
'            the appropriate parameters.
'       6.71.37  Optimized code for fast target pressure acquisition for Becton-Dickinson. TAR040622
'       6.71.36) 04 06 06 Tim - Added Mettler Balance in place of penetrometer for
'               Becton-Dickinson
'       6.71.35) 5-24-04 MJLC
'        1) Updated language files in "About", main form, manual control, pressure step list, and
'            autocal.
'       6.71.34) 5-19-04 MJLC
'        1) Localized a label in calib_reg that I'd overlooked.
'       6.71.33) 5-14-04 Ron
'           1) All pleasewait.hide changed to unload pleasewait.  If you were running qc mode and
'               ran a test it would leave the pleasewait form loaded but hidden and quiting the
'               program from the qc main form did not unload it so the program stayed on but hidden.
'              Also changed lv_man_ctrl from hide to unload and then commented it out because
'               by that point in the program the form had already been unloaded.
'           2) Leave venting valve open at the end of the pressure hold test.  On some systems
'               the vent wouldn't have been left open long enough to properly vent the entire
'               system, so leaving the venting valve open until they try to run something else
'               should be better.
'           3) Modified bubble point routine so it won't trigger the bubble point if the pressure
'               drop was caused by a change in the range of the pressure gauge.
'       6.71.32) 04 05 14 Search for Tim Richards
'           added global boolean g_bBalanceNotPenet and the subroutine ReadBalanceNotPenet.
'           This subroutine is called by ReadXReturnX4 after which ReadXReturnX4 immediately quits and
'           passes back the value. The global is set by reading the H2OPerm ini switch, set to B,
'           along with feature 16 set (meaning liquid permeametry).
'       6.71.31) 5-11-04 Ron
'        Fixed timer so that it won't lock up on requested delay of 0 seconds.  This showed up when you
'         tried to run a leak test with the read delay parameter set to 0.
'       6.71.30) 5-11-04 MJLC
'        1) Further updated the language files at the request of Mr. Yaza.
'       6.71.29) 04 05 10 Tim
'           line 15114 modifying the report writing code initially to work with Heidi's ambient/elevated
'           liquid perm data files. Procedure RunTest(). Series of Select/Case. This code is written
'           line-wise, not report-wise, which makes it difficult to add a line into a single report -
'           difficult to parse.
'       6.71.28) 4-29-04 Ron
'        1) Added support for digital door switch in compression systems.  Autopiston systems always had the
'          door switch of this type.  Recirculation systems alway have a door switch of a different type.
'          Both autopiston and compression can have the safety keypress system.
'          Now, setting safetyup or safetydown to "A" (it used to be "Y" or "N") means that instead of a
'          key press it now requires the door switch.  The safety keypress module now can release on the door
'          switch or the key press depending on the "Y" or "A".
'        2) Changed low flow calibration routine so it applies temporary higher flow target at beginning if the
'          low flow controller is not responding.  This should make it work with newer flow controllers that
'          are slow to respond at the very beginning.
'       6.71.27) 4-27-04 Ron
'        Added support for microflow gas permeability, just like microflow porometry
'       6.71.26) 4-27-04 MJLC
'        1) Updated strings in language file to incorporate recent program additions.
'       6.71.25) 4-26-04 Ron
'        Changed waiting and rs232 communication so that it releases time to the cpu whenever it is waiting.
'        Many time delays changed to calls to existing waitseconds routine.
'        waitseconds changed to use new waitms routine (in wait.bas module)
'        waitms2 routine added to wait.bas module to handle second comm port (for watlow)
'       6.71.24) 4-23-04 Ron
'        Added in features from branched version 6.71.22a and 6.71.22b
'        1) (from 22b) Tuned up pressure regulator switchin in the middle of bubble point to make it work better
'        2) (from 22a) if Compression_Increase_Factor is negative, limit the maximum pressure for
'        the test to the compression pressure divided by the absolute value of the Compression_Increase_Factor.
'        (If the Compression_Increase_Factor is positive, the compression pressure is increased as the testing
'        pressure goes up.  If it is zero, the compression pressure remains a constant and there is no limit
'        to the maximum pressure of the test.)
'       6.71.23b) 4-15-04 MJLC
'        1) Automatic calibration of pressure gauges, too!
'       6.71.23a) 4-9-04 MJLC
'        1) Changed the "Autocal" form so that calibration of flow meters is, indeed, automatic.
'       6.71.23) 3-29-04 MJLC
'        1) Rearranged items in TitleScrn.Load() so that the "languages" folder isn't created in
'         a code directory. (If by chance capwin.ini isn't found, the error message will come up in
'         English, not the user's preferred language.)
'       6.71.22) 3-25-04 Ron - added in more pressure regulator switching
'       6.71.21) MJLC 3-24-04  No new changes -- this is a version of the code integrating Ron's v6.71.20 with Matt's
'        v. 6.71.19a-d.
'       6.71.20) 3-5-04 Ron - all changed commented with 6.71.20
'        1) Added support for switching high flow meters without moving valve 10.  This happens when
'         the variable "suspend_v10" is true.  New variable switch_high_flow_enabled must be true.
'        2) Added capwin parameter second_regulator_starting_point, which is the number of counts of
'         the second regulator that equals the 4000 count point of the first regulator.  This defaults
'         to 0, which suspends its usage
'        3) The learning routine for deltap=0 liquid perm tests is now disabled for pressures less than
'         25 PSI.  It wasn't working for really low pressures, and the customer who wants this is only
'         using pressures of about 70 PSI.  If there is customer demand, maybe this will be made to work
'         on low pressure tests as well.
'        4) Added trap for bad value for penetrometer_start_test_point - if it is lower than 50% of the
'         average of pen500 and pen20500, it will be reset to pen500.
'       6.71.19d) 3-22-04 MJLC
'        1) Added a custom test type for Saint-Gobain. It allows them to set the regulator and v2,
'          clamp a test fixture down onto a sample, wait a few seconds, and take a reading. For now,
'          since this is in the developmental stages, this is activated by a button in Autotest marked
'          "Saint Gobain". This button is hidden unless sending a new executable to SG.
'           More modifications to CVCalc to get all the checks working with the new calculation.
'       6.71.19c) 3-18-04 MJLC
'        1) Modified lohm back-correction to use the new formula Ron calculated.
'       6.71.19b) 3-17-04 MJLC
'        1) Modified openv2completely to trap possible error reading open limit in Sam Bo/SK machine.
'       6.71.19a) 3-16-04 MJLC
'        1) Fine-tuned pressure step list routine to work better in bar for Sam Bo
'       6.71.19) 2-24-04 MJLC
'        1) Fixed a bug in the pressure step list interface; list is now sorted properly as values are
'         entered.
'        2) Moved the "Change Supervisor Password" menu item to the Mode menu (under Group).
'        3) Added a Debug menu to TitleScrn. This is intended to give easy access to the various debugging
'         variables without having to change the ini file. The menu is hidden to the user unless
'         debugMenuVisible=Y. Initial values are taken from the ini file. For now, checking or unchecking
'         an item will make it valid while capwin is open, but changes are not saved. The "stability debug"
'         option is a little different: toggling it does not turn on stability_debug, but makes the original
'         menu item in the Progress screen visible.
'       6.71.18) 2-14-04 Ron
'        1) Added interpolation of the data file at the end of the gas permeability test if you are
'         using the pressure step routine.  The original raw data is stored to lastdata.cft.  The
'         interpolated data is stored to lastdata2.cft and then copied to where they wanted it.
'       6.71.17)  1-23-04 Ron
'        1) Added support for external Watlow connected to second serial port of computer.  You have
'          to edit the capwin.ini file to set this up and enable it.
'        2) If the external Watlow is enabled and working, collect temperature data from it during the
'          gas permeability test, and use the "TEMPERATURESENSOR" header on the data file to signify
'          that there is one set of temperature data present.  Note that we haven't used this header
'          in such a long time that the current report program does not support it.  (It supports the
'          "TEMPERATURESENSOR2" header that signifies two temperature sensors for liquid permeability
'          tests.
'        3) Added support for "TEMPERATURESENSOR" header in common routines for loading and saving of
'          data files.
'        4) Added support for new valve 23, for liquid sample chamber drain.  Currently it is only
'          supported for recirculation systems, and it is closed before recirc1 is entered and only
'          opened manually when you are able to open the sample chamber.  You open it first, to drain
'          out the excess liquid before you open the sample chamber.
'       6.71.16)  1-12-04 Ron
'        1) Removed unload of piston in load sample routine if you just did a vacuum purge or if
'          you have the load prompts turned off.  This caused a problem for GE where the piston
'          would unload and then load again and this could cause the piston to pop up and then back
'          down before their test started.
'       6.71.15)  1-9-04 Ron
'        1) Added "learning" of pressure overshoot during single pressure elevated pressure liquid
'          permeability (deltap=0).  It remembers the overshoot from the pressure target (if any) and
'          applies this to the next deltap=0 test as long as the target pressure remains the same as
'          the last test.
'        2) Also fixed problem where if you only have one defined group (apart from the default group)
'          and you are currently in that defined group and tell the program to delete groups it would
'          crash.  (This is because you can't delete the default group or the current group, so the
'          list of possible groups you can delete was empty and it didn't like this.)  Now, in this
'          situation, it will put up the empty box saying there are no groups you can delete, just as
'          if there was only the default group.
'       6.71.14)  1-2-04   MJLC
'        1) Made pressure step list follow user's choice of pressure units -- previously was assuming
'          PSI only.
'       6.71.13)  12-18-03 MJLC
'        1) Found that the code initializing the pressure step list was missing in the latest version,
'          causing an error. Added it back in.
'       6.71.12)  12-9-03 Ron
'        1) Added "pause" button to lohm calibration status form.  This allows the operator to pause
'          the lohm calibration to let the compressor catch up with the pressure.
'       6.71.11)  11-26-03 MJLC
'        1) Modified "pass/fail" option in select_test so that it's only shown when valid (i.e. for
'          capflow and BP.
'        2) Fixed a bug in the changing of length units in preferences.
'        3) Added support for setting a pressure list for a capflow or gasperm test. Use PS_usingList to
'          determine if the option is being used. Lists are created from the "pressure list" menu option
'          under "Modify" and specified for a test in the select_Test window. Control code is in run_c_pass
'          starting in section 3251 and 3300.
'       6.71.10a) 11-17-03 changes made in Japan, may need to be merged with other changes later
'        1a) Changed default minimum_liquid_test_stop_point to the top of the penetrometer.
'          It had been 50% on the penetrometer and if you are running a non-porous sample
'          you will never get this low and the test will never end.
'        2b) Added "charset" setting in set_fontname (it was already in set_fontstuff) so the
'          "pleasewait" form shows properly in Japanese.
'        3c) Fixed problem where if you have a valid SCDiam (from a previous ESA test), it would
'          be stored in subsequent gas or liquid permeability test files, which the report program
'          can't handle.
'        4c) In liquid permeability, if the starting pressure is 0 and the penetrometer is above the
'          the penetrometer_start_test_point, the regulator will be increased slowly, up to a maximum
'          of 2 PSI above atmospheric pressure, until the penetrometer gets in range, and then the
'          regulator will be zeroed and the penetrometer vented before the test is actually started.
'        5d) Made some more variables double (for pressure storage) so they appear correct when shown
'          back to the user.
'        6d) Added capwin parameter of Compression_Increase_Factor, which defaults to 0 to disable the
'          new effect.  If you are running with a positive compression pressure, the compression pressure
'          will be increased as you run the test so that the actual compression pressure is always at
'          least the compression pressure you asked for plus the Compression_Increase_Factor multiplied
'          by the current testing pressure.  A good Compression_Increase_Factor for normal pistons is 2.
'          This was discovered experimentally.  This will stop the sample chamber from opening up when the
'          force trying to open the sample chamber due to the testing pressure exceeds the compression force
'          trying to keep it closed.
'        7e) Use Compression_Increase_Factor in the elevated pressure liquid permeability test as well.
'        8e) Changed regulator calibration so it takes more points at the beginning of the curve.
'       6.71.09) 11-13-03
'        1) Added units to some of the extended information added to the data file
'       6.71.08) 10-27-03
'        1) Fixed problem introduced in version 6.71.06 when writing microflow settling
'          parameters, this would cause liquid permeability tests with temperature information
'          to get messed up during the final file copy.
'       6.71.07) 10-17-03
'        1) Changed name of "recirculation 1" and "recirculation 2" to "Partial Recirculation"
'          and "Full Recirculation" at the request of Ballard.
'        2) New parameter in the capwin.ini file: minimum_liquid_test_stop_point
'          This defaults to the 50% point of the penetrometer.  At the end of the liquid
'          test, if you run out of data points or maximum pressure, the test will continue
'          until the penetrometer reaches this point or below before terminating.  This is
'          so that the penetrometer is low enough so that when you open the sample chamber
'          there will be room for the liquid in the chamber to go back into the penetrometer
'          and you don't build up too much liquid so it overflows when you open the chamber.
'       6.71.06) 10-3-03
'        1) Vacuum Purge is now available for those systems with "vacuum_purge_enable"
'        2) Microflow Settling parameters are stored to the data file, along with the final
'          settling time.  This requires that the data file be read in and re-written after
'          the test is over because the total settling time is stored in the header of the
'          file, and this is written before the test is started and the total settling time
'          is not determined until the test starts.  If there is anything else about the data
'          that needs to be changed after the test, this will be a good place to put it.
'        3) Changed the NOWHERE counter so it lets the test go longer with no data in attempt
'          to stop special regulator-only tests from timing out before any real data is
'          collected when the regulator has a large zero offset.
'        4) Moved compression pressure storage to user ini file so each user could have his
'          own compression pressure.  Piston area is still a system-wide variable since no
'          matter how many users, they will all still use the same piston.
'        5) Test changes for Seika Toyota - valve setup sequence for air-bottom hydrohead
'          with lower penetrometer
'        6) Changed elevated liquid perm so first part doesn't wait for pressure increase
'          if you are running deltap=0 (for ballard)
'       6.71.05) 9-30-03
'        1) Fixed problem where recirculation systems, when running tests other than liquid
'          permeability, can get a "subscript" error at then end of the test.  This is due
'          to the temperature array being read when it had not been filled.
'       6.71.04) 9-29-03
'        1) Fixed problem introduced on around 9/8/03.  On systems without microflow, if you
'          enter the test type selection form it will improperly enable the microflow test
'          type.  If you then run any type of test and enter manual control from within the
'          test, it will get an overflow error when it tries to read the microflow pressure
'          gauge.
'       6.71.03) 9-22-03
'        1) Cleared out some counters when you exit from hold during the bubble point routine
'          to try to fix problem with overflow reported by Seika in Japan.
'        2) Fixed test setup so if you change the preferences for seal length while you
'          are in the setup screen it will correctly show the microflow test parameters.
'        3) Split variables for hold time and pressure for pressure hold test and microflow
'          They used to interfere with one another if you switched between the two types of
'          tests.  Also, the preferences window allows setting of the hold time for the
'          pressure hold test, and this could mess up the microflow test.
'       6.71.02) 9-16-03
'        1) For recirculation systems, added pulse of fill valve open and close at end of test
'          to relieve any built-up pressure back into the fill tank before opening the sample
'          chamber.  Seika reported that with some samples they would have a blast of liquid
'          when opening the sample chamber at the end of the test.
'        2) Added temporary output file during recirculation liquid permeability.  This will
'          eventually be added to standard data file.
'        3) In elevated pressure liquid permeability, if the penetrometer is still above the
'          penetrometer_start_test_point, the main routine will not take any data until the
'          penetrometer has dropped down enough.
'        4) Program no longer crashes if you try to delete users when there are no users
'          other than the default user (which you can't delete).  This problem showed up
'          when the user list became sorted.
'        5) On systems with the new door lock, it now prompts them to open the door in the
'          middle of two-pass tests (because they need to be able to open the sample chamber
'          to wet the sample and reinstall it or put in a dry sample.
'       6.71.01) MJLC 9-15-03
'        1) Added a shortcut to the preferences window from the test setup screen.
'        2) Added a checkbox for the "linear seal" microflow option on the test selection screen.
'        3) Modified the linear seal functionality: added an indicator variable (seal_state) to
'          differentiate between a simple entered seal diameter and the case where cyl_len > 0,
'          where we need to enter both an inner and outer diameter. Set up do_final_copy to print
'          inner and outer diameter if they are non-zero; calculations for effective diameter
'          will be done in caprep.
'        4) Added support for recording temperature directly into a microflow file. This is
'          indicated by the special filetype "DIFFPERM+t" and flagged in the software by MF_recordTemperature.
'       6.71) Released 8-29-03
'       6.70.61)
'        1) End of high pressure liquid permeability test changed for airtop machines to minimize
'          getting fluid back into the pressure manifold.
'        2) Added support for auxin bit in version 7 feature number (=1024)
'          Manual control screen can now handle both compression pressure and temperature readings
'          Whichever was the last one you clicked on will be the one displayed (since they both use
'          the same line on the display)
'        3) Prompt for step pressure moved to before prompt for maximum pressure in elevated liquid
'          perm test, so if step pressure is 0 it won't bother to prompt for the maximum pressure.
'          If the step pressure is 0, the test will automatically run to the end of the penetrometer
'          or to the maximum number of data points (whichever hits first) and then stop.  The maximum
'          pressure will be ignored.
'        4) Door switch added for recirculation test.  If door is opened, the test will abort.  Also
'          you can't go into recirc2 with the door open, and if you are in recirc2 and open the door it
'          will go back into recirc1.
'        5) Added optional delay to wait for stable pressure before microflow pressure test.
'        6) Changed order of valves for liquid perm test to help avoid getting liquid in the pressure
'          side.
'        7) Added support for "doorlock" boolean - if you have a door lock and need to use a special
'          command to open it back up.
'        8) Moved compression pressure entry to beginning of test setup so it will be done before the
'          header is written to the data file so we can write the compression information to the data
'          file.
'        9) If you have temperature reading, the main testing routine will log temperature and pressure
'          and flow to a "tempdata1.txt" or "tempdata2.txt" file in the main directory.  Wet curve will
'          be in "tempdata1.txt", and dry curve will be in "tempdata2.txt".  Microflow data is still in
'          "tempdata.txt".  This will be overwritten by the next test.  If you have two temperature
'          probes, the second one will be used for the tempdata file (as the first one is assumed to be
'          for the liquid permeability chamber).  Also, recirculation systems do not zero the liquid
'          temperature at the beginning and end of the test, but leave it at whatever you had it set to.
'       6.70.60)
'        1) Added support for "recirculation=Y" - recirculation pump and two more valves for liquid perm
'          as used on the Ballard machine.
'        2) Now zeroes compression regulator (electronics version) before setting the compression
'          pressure.  It had left the pressure alone, which could result in too high a compression
'          pressure.  This was discovered on the GE machine.  This doesn't affect motorized compression
'          regulators like on Ballard.
'        3) Doesn't ask for seal diameter if you don't have the microflow seal length option turned on.
'          and doesn't save the seal length to the data file unless the option is turned on and the
'          diameter is greater than 0.
'       6.70.59)
'        1) Added a new function, File_Exists, and used it to check for the presence of the lohm
'           file specified in user preferences when the software is loading. (load_user_global_stuff)
'       6.70.58) -- MJLC 7/30/03
'        1) Moved the "microflow porometry" option in select_test because it was covering up "square root" calc dry. (select_test)
'        2) Fixed a bug where "use min pressure in dry curve" pref. wasn't being saved properly. (prefs form)
'        3) Changed Load_Sample so that we don't have to wait for flow meter stability when
'           doing a pressure hold test.
'        4) Enabled launching caprep at end of test for BP, LV perm, and press. hold tests. (start_caprep variable)
'        5) Changed pressure hold test so that graph does not scale out automatically if not checking dp/dt.
'        6) Updated language file. (loadtextstrings in each form and ts variable definitions at top)
'        7) Fixed a bug in the pressure hold test where pressure drop was not calculated correctly
'          for tests measuring PSI/min.  (in pressure_hold)
'        8) Added P to the bubble point log. (BPdebuglog.txt)
'       6.70.57)
'        1) Improved the autocal interface: "stop" button works, user can set maximum pressure/flow,
'          final report shows correct percentages, better regulation.           -- MJLC 7/3/03
'       6.70.56)
'        1) New autocal routine for calibration of pressure gauges and flow meters. Accessible via manual control.
'          Still pretty buggy.                                                  -- MJLC 6/20/03
'       6.70.55)
'        1) Partial update to language files.                                   -- MJLC 6/17/03
'        2) Fixed a bug in the looping GP and BP tests, although there may still be a problem running
'          a looping BP test when "auto increment filenames" is checked.        -- MJLC 6/17/03
'       6.70.54)
'        1) Added an option to input a "seal length" for microflow tests so that flow can be calculated
'          as passing through a linear length instead of a cross-sectional area.        -- MJLC 6/16/03
'       6.70.53)
'        1) ADC calibration form now correctly resizes for version 6 machines.  It had resized to
'          hide those parts that are only used by version 7, but then the form was expanded
'          and it then resized too much and the exit button was partially hidden.
'        2) Now allows use of motorized compression regulator on version 7 porometers by setting
'          motorized_compression_regulator=Y.
'        3) two-stage filling now works for any bottom-up autofill liquid permeability
'          test, not just for liquid perm only machines.
'        4) New variable max_fill_point, which defaults to the pen500 value.  This is where the
'          autofill valve will close during filling for liquid permeability.
'        5) New option under preferences for liquid permeability for compression and autopiston
'          machines - you can select to delay the compression until after the initial fill or
'          do the compression before the initial fill.
'        6) New variable piston_area gives the compression piston cross sectional area.  This
'          defaults to 1 square inch.  This, coupled with the user's entry of the sample cross
'          sectional area, is used to modify the compression pressure value that is entered
'          into a value for compressive force on the sample.
'       6.70.52)
'        1) Microflow volume calibration now uses a lower pressure for the flow controller
'          because some flow controllers would leak at high pressure.  It now uses the the
'          max_bp_pres_dif plus 14.7 for the initial pressure so it guarantees that it can
'          keep flowing all the way up to 14.7 above atmospheric pressure.
'       6.70.51)
'        1) After the penetrometer reaches the penetrometer_start_test_point, the air pressure
'          that was used to move it to this point is released so we start at as low a pressure
'          as possible.
'        2) Target pressure for elevates pressure liquid permeability now takes into account
'          the liquid level height.  This mainly affects low pressure testing.
'        3) min and max pressure boxes only appear if you are doing a test that uses these
'          values.  (liquid perm, pressure hold, microflow so not use these)
'        4) tortuosity factor box only appears if you are doing a test that uses this.
'        5) Improved selection of dual regulator for liquid perm, pressure hold, and microflow
'          so selection is based on maximum pressure for these tests, which are entered
'          differently, not the maximum pressure from the parameter file or test setup
'          screen (since this is now hidden for these types of tests).
'       6.70.50)
'        1) Initial autofill of penetrometer for bottom-up machines has been improved.
'          New variable bottom_fill_point.  The isolation valve (valve 20) will be closed
'          when this point is reached during fill, and then the penetrometer will be
'          filled up to the top and then valve 20 will be opened.  This should fill the
'          sample better and top off any space above the sample when using adapter plates.
'          New variable sample_zero_point gives the height on the penetrometer where
'          the sample is actually at zero.
'          When the penetrometer is refilled, it will also use this dual-stage filling method.
'          A new button in manual control will allow you to set the sample_zero_point and
'          the bottom_fill_point.  These may be removed at some later time when an automated
'          method for determining these is developped.
'          A new preference for liquid perm is added - norefill.  If true, the elevated pressure
'          test will not attempt to refill the penetrometer - the test will end if a single
'          pass of the penetrometer is not enough to get all the data points we asked for.
'        2) After initial fill, the regulator is increased slowly to make the penetrometer go down
'          until it reaches the penetrometer_start_test_point (in cm).
'       6.70.49)
'        1) Initial pressure regulator increase is removed if you have a solenoid valve 2.
'          This stops the first data point from being too high.
'       6.70.48)
'        1) Moved diffpg reading in capstuff reading to above liqpermonly initialization.
'          if you have a diffpg and no high flow meter, you are now a microflow porometer
'          and liqpermonly is not turned on.
'        2) OpenV2Pos now opens solenoid valve if there is one
'        3) Changes to diffperm test at end when lowering pressure slowly
'        4) Changes to find_volume to give some flow when initializing flow controller
'          Reversed flow controller doesn't go up just because you increase the pressure on it
'        5) microflow test now allows microflow pressure to go down and still save the data.
'          The report program will have to accept this type of data file.  Also, the back pressure
'          correction and pressure regulator increase now uses smaller steps because it was over
'          correcting.
'        6) New option to turn off regulator control during microflow test.  This is on the new
'          microflow tab of the preferences screen.
'        7) new file gasflowconversion.ini will be created if it is not there.  This will default to
'          air/nitrogen as the only gas, with a conversion of 1.  If there is only one gas, the program
'          will look the same as before.  If there is more than one gas, then an extra item will
'          be on the main screen showing what gas is selected and you can use this to select a new
'          gas.  The gas conversion factor will be used for all flow readings and parameters.
'          The selection box has been modified to allow selection of gasses as well as group names.
'          In doing this, the selection box has been updated so it preselects the existing group or
'          gas name, except when deleteing a group in which case nothing is selected and the current
'          group is not displayed (since you can't delete the current group).
'        8) Piston is raised at end of microflow test.
'        9) Slow vent at end of microflow test gives more information and delays a little more at the
'          end to let the regulator go all the way down to 0 before venting the system fully.
'       6.70.47)
'        Added a new code module, Statistics.bas, which contains functions for calculating mean and
'         linear regression for data in a new structure, xy_data. Used this to add linear regression
'         as an option to the pressure hold test (when determining pressure drop at the end of the test).
'        Now storing the chamber volume, Chamber_Volume, in the .ini file during the microflow calibration.
'        Added a new option under the Calibrate menu for doing this if you don't have microflow.
'        Added volume leak information to the pressure hold test results, and also to the data file.
'        Removed P02 (P0 squared), which was useless.
'       6.70.46)
'        Safetydown variable takes priority over the door switch.  Autopiston used to assume a door switch, and the
'         safetydown variable was ignored.  Now, if you have the safetydown turned on, it will ignore the door
'         switch (if any) and use the safetykeypress form whenever you want to clamp the sample chamber.
'       6.70.45)
'        Replaced Show_Result and start_caprep variables with auto_report_type (but left them in for version
'         compatibility). User can now choose to automatically run a report at the end of a test, or just run the
'         summary sheet or equivalent. Per KG's request, the old "show results at end of test" option is now hidden
'         because the results do not agree with caprep's calculations.
'        Corrected a bug in the "fail by absolute pressure change" pressure hold option.
'        Changed the format of the BP debug log.                                    -- MJLC 4/23/03
'       6.70.44)
'        Minor corrections to the writing of the pressure hold test data file.      -- MJLC 4/17/03
'       6.70.43)
'        Replaced the call to automatically start Caprep at the end of a test, which disappeared under
'         mysterious circumstances.  Also fixed a bug which caused a crash when changing the liquid name for a
'         permeability test. -- MJLC 4/15/03
'       6.70.42)
'        Added a debug log to the bubble point routine - activated by debugBP=Y in capwin.ini -- MJLC 4/10/03
'        Fixed a bug created in 6.70.37 where line26 was not updated properly and BP routine misbehaved for
'         machines with a low flow controller.
'       6.70.41)
'        Modified the pressure hold test to save hold rate unit information (PSI, PSI/sec, PSI/min, etc)
'         at the end of the data file.
'        Pass/fail test: changed the background color on the label in the progress screen for better contrast.
'         Added the coveted smiley and frowny faces. Also moved the pass/fail check slightly so that it fits
'         better into the main BP loop.                 -- MJLC 4/2/03 (one whole year at PMI...)
'       6.70.40)
'        Microflow test now lowers pressure slowly so that there is no shock to the sample.
'         The internal pressure is lowered to keep it no more than the initial holding pressure
'         below the microflow pressure to allow the pressure in the microflow pressure volume to
'         go back through the sample.  Once the microflow pressure is less than the initial
'         hold pressure, it is save to zero the pressure in the chamber and then vent the system.
'         If you are using a high pressure for your test, this will have no effect.  This will
'         only do something if you are running a low pressure (like 1 PSI) test and the microflow
'         pressure chamber builds up beyond 1 PSI (which would make the forward pressure above 2
'         PSI).
'        Removed check for flow meter off scale during main equilibrium routine - this will allow
'         low flow samples that have high flow rates initially after every pressure increment to
'         have time for the flow to go back down.  If the stable flow is still above the maximum
'         flow rate then the test will stop.  Note that you will have to increase the equilibrium
'         iterations to a large number to give the system more time for the flow to drop.
'       6.70.39)
'        Added current lohm table to parameters saved in the data file.
'        Fixed a bug in the pressure hold test converting /sec to /min.     -- MJLC 3/26/03
'       6.70.38)
'        For reversed flow controllers, the pressure regulator now increases as the bubble
'         point test goes along so that we keep a fairly constant delta-p accross the
'         flow controller.
'        The maximum bubble point differential pressure is now derated depending on the
'         atmospheric pressure.  At two atmospheres (absolute pressure) the bubble point
'         differential pressure will be divided by 2.
'       6.70.37)
'        Fixed flow controller pressure increase routine so it doesn't trigger just due to
'         noisy flow readings - the flow reading must stay below the target for 10 seconds
'         before it will try to increase the pressure.
'       6.70.36)
'        Removed some test code that accidentally got saved into the program. -- MJLC 3/20/03
'       6.70.35)
'        Internal pressure during microflow test is now corrected for the back pressure
'        If internal pressure falls too low, the regulator will increase to compensate.
'       6.70.34)
'        Made a few minor adjustments to the elevated liquid perm routine.  -- MJLC 3/13/03
'       6.70.33)
'        Now always writes the tortuosity value to the data file during a test (previously only if not default).
'        Set the font properties of the entire manual control and testscrn forms to alleviate the scrolling/wrapping problem.
'        Added a new user preference for the minimum time between data points in liquid perm (previously hard-coded to
'         0.2 seconds).
'        In the elevated liqperm test, added the mintime criterion and decreased the min. step from 1/20 DAC to 2.5%  -- MJLC 3/10/03
'       6.70.32)
'        Fiddled with the new liquid perm routine a bit more; added the new timer to the elevated
'         perm routine.                                 -- MJLC 2/28/03
'       6.70.31)
'        The zerotime parameter is now always set to 10 seconds for version 7 with
'         low flow controllers.  This parameter was previously not used in version 7
'         (it was used in version 6 and before) and then was used again to control the
'         timing of the low flow controller reset procedure.  The default value of 1
'         second (for version 6 and before) was too small for this new use in version
'         7 and we don't want to have to make everyone change this parameter for version
'         7 so it is now hard-coded at 10 seconds and the parameter that the user can change
'         is now ignored.  Users who had set this to 10 seconds should set it back to the
'         default of 1 second in case it gets used for something in the future.
'        Also: Internal pressure during microflow test is now corrected for the back pressure
'       6.70.30)
'        Continued improvements to the pressure hold test. Fixed a problem with the pass/fail calculation and improved
'         the seconds/minutes selection.
'        New method of doing liquid perm test:
'         added Ron's new implementation of the "performance counter" method of timing -- support for QueryPerformanceCounter
'         and QueryPerformanceFrequency in "kernel32" dll; also added Ron's new function, time_difference, for finding the time
'         in seconds between two consecutive calls to QueryPerformanceCounter. These are now used for recording the critical time at which
'         the height and pressure are recorded in the liquid perm routine.
'        Corrected a bug in the translation strings where "low pressure gauge" was mapped to "low flow"
'       6.70.29a) (Special version)
'        Includes first 5 changes in 6.70.29. Changed time to wait between points in
'         liq. perm tests from 0.2 to 1 s; disabled press. hold in Preferences tab since it is not yet completed.
'       6.70.29)
'        More change requests from Haemonetics:  Ripped apart QC mode and redid the interface.
'        Pressure hold tests: Added option to fail based on absolute pressure change instead of pressure/time.
'        Redid the averaging method used: now takes as many readings as user wants and averages them to obtain a single point.
'        Removed the subroutine Hold_Key(), since all it did was call DoEvents.
'        Added a debugging log for elev. liq. perm, set by debugh20perm in capwin.ini. - saves pressure and height
'        Rearranged the preferences for the pressure hold test and pulled some options from the test setup screen to the preferences
'         window.
'        In the liquid perm test, changed the minimum time between samples from 0.2 to 1 sec to reduce noise in the height readings. May
'         want to make this variable later.
'       6.70.28)
'        Added support for "reverse_flow_controller".
'        Added min and max logging of pressure gauge in manual control for debugging purposes
'       6.70.27)
'        Microflow porometry test now subtracts microflow pressure (back pressure on sample) from internal
'         pressure so that we use the proper differential pressure for the test.
'       6.70.26)
'        Discovered that the data editor window was much too large to fit on an 800x600 screen! Dunno
'         how I missed that one. Managed to squeeze it in by rearranging everything on it.
'        Reworked the pressure hold test again at the behest of Dr. Gupta and the Haemonetics people. Sampling rate is
'         now variable. Progress screen shows and records averaged data, not raw data. Test does not automatically fail
'         if pressure falls too much; instead, it's linked in to the pass/fail option in the test selection window.
'       6.70.25)
'        Microflow calibration for air-top machines is now working.
'        Pressure hold test now uses regulator calibration table (for new machines) to set the initial
'         pressure faster.
'        There is now a delay after the minimum bubble point pressure is reached before the bubble point
'         test starts to let things stabilize better.
'        There is now a delay after the first pass before the second pass starts if they are not going to
'         be prompted to load the sample.
'        Hid the "Set Valve Limits" menu option.
'        Reworked the reworking of the lohm table setup. A lohm table pathname is now stored for each user. When
'         the program starts or a different user is selected, the lohm table at that pathname is copied into the
'         lohmtable.cal file in the main directory.  When a new lohm table is created, it is saved to the user's choice
'         of filenames.  The lohmtable.cal file in the main directory is only a working file for the software and is overwritten
'         each time the program starts up or the group is changed.
'       6.70.24)
'        Removed procedure update_board_loc, since it was a one-liner only called from one place.
'        Fixed a bug in display of comm ports in preferences window.
'        Fixed a problem in the autoparm print routine caused by unpleasantly long filenames.
'        Added checking for lohmtable.cal in capwin directory as well as user directories.      -- MJLC, 1/14/03
'        Made the preferences window non-resizable.
'        Disabled the section that would skip data points during the Lohm calibration if the readings were
'         not stabilizing fast enough.  This was messing up on new faster version 7 instruments and skipping
'         points that really should be taken.
'       6.70.23)
'        Changed the run_elev_lqperm routine to handle problem with zero pressure.  (Ron)
'        Enabled font changing and translation for a couple of lines in the manual control screen that I missed. Updated
'         the language file with the new windows and text that have been added recently.
'        Modified the pressure hold test to include a user-changeable averaging method (variable num_PH_AvePoints). Also modified
'        the press.hold test display, changing the y-axis to scale out as the test progresses and adding two lines to indicate the edges
'        of the "pass/no pass" region.
'        Fixed a bug in opening the leak test file at the end of the test
'        Changed the location of lohmtable.cal files so that they are now group-specific. Haven't changed the default directory
'         for backup tables, though; still goes to \parms.  Added display of the current lohm table to the title screen                -- MJLC 1/10/03
'       6.70.22)
'        Enabled automatic starting of CapRep at the end of a test (finally!). Modified the QC window to make it harder to
'         select the wrong test.                                        -- MJLC 1/3/03
'       6.70.21)
'        Bubble point test now starts at pressure differential specified by new capwin.ini parameter max_bp_pres_dif, which
'         defaults to 20 PSI.  Also uses new low flow controller initialization.
'       6.70.20)
'        For dual regulator systems, both pressure regulator calibration tables are now loaded at all times.  The proper table
'         is used depending on which regulator is active.
'        The proper regulator should be selected based on the maximum pressure for the test as set up on the test setup screen
'       6.70.19)
'        Removed all code references to version 5.x hardware, the small inst. board form, win95io.dll, and Version 5-only subs.
'        Made substantial changes to TitleScrn menus by collecting various options into a new window, prefsForm. Removed
'         the setuplogging window, since the options are now in "prefs". Hopefully this will make things simpler
'         for the user and present a cleaner-looking interface. Also updated the PMI logo in the main screen and removed the logo from cap_cur.frm.
'        Also removed duplicate copies of update_linear_unit and update_thickness_unit functions from TitleScrn.
'        Also added the ability to start CapRep with the current data file at the end of a test -- currently disabled because file does not load properly.
'        Removed initialize_psr from capflow.bas, since it was a one-line function that was called exactly once.
'        Modified the lohm calibration routine so that users can select the starting flow point (% of max. flow).
'        Fixed a problem with the hidden label in the test screen - font wasn't getting changed along with the rest of the program, so
'         long filenames weren't being displayed properly.  -- MJLC, 12/02
'        Added in error trap in serial port error correction routine to stop possible
'         endless loop if you get a specific type of communications error while reading
'         the penetrometer.
'        Added a routine to calibrate both regulators in a dual-regulator system when "calibrate regulator" is selected ... but it's currently
'         disabled because it's not working properly yet.
'        Added an option for the pressure hold test fail rate to be displayed in PU$/min instead of PU$/sec.
'       6.70.18)
'        Added the capability to run looping tests for demonstrations. This is accessed by setting LoopingDemo=Y in
'         capwin.ini and selecting the test type as usual in the software. There is no way for a user to set the
'         LoopingDemo variable in the software. Also removed an undesirable "Yo!" message that I inadvertently
'         left in the code.                         -- MJLC, 11/25/02
'       6.70.17)
'        Stopped regulator from going too high at the end of the bubble point before
'         the wet curve starts.  This could be caused if the regulator has a large
'         zero offset, which would be doubled in the old program.  If the regulator
'         if too high, the pressure would go up too fast when valve 2 was opened and
'         you could skip some data points.
'        Also removed a 0.6 second delay (two 0.3 second delays in a row) from the start
'         of the test that were not needed.
'       6.70.16)
'        Modified the leak test filename format to include date and time, so that multiple tests
'         will not overwrite. Also changed the Lohm calibration to start at a higher flow rate.      -- MJLC, 11/1/02
'       6.70.15)
'        Added debug button to progress screen that will bring up a small window that
'         will show information about how the stability routines are doing.
'       6.70.14)
'        Removed units of measurement from the translated .ini file. Also added a pass/fail
'         test option based on bubble point -- parameters are in select_test. Also added a variable
'         delay in the leak test routine before reading the initial pressure at each step. This should
'         give the gauge time to stabilize before taking the reading. Set by read_delay in autoparm. -- MJLC 10/30/02
'       6.70.13)
'        Added support for second regulator calibration file.  When using reg 2,
'         the calibration file is now capwinrg2.cal.  This means you don't have to
'         recalibrate when you switch regulators.
'        Also fixed problem with i/p regulator at beginning of dry curve where it
'         sometimes would not increase the regulator enough to get any flow until
'         valve 2 was all the way open.
'       6.70.12)
'        Added storage of last 100 transactions over the serial port to the debug log
'        Also added checking of return values from version 7 machines to make sure
'         they are in the right range so that an error character does not cause an
'         overflow or other numeric or procedure error.
'        Added error trapping to raw reading procedure to try to trap a very rare
'         procedure call error.
'        Fixed error in error trapping routine (!) that would cause the very rare
'         procedure call error (left in the trap above).
'       6.70.11)
'        Changed comm timout values to try to make system more stable on computers with slower
'         comm ports.
'       6.70.10)
'        Added log_comm to capwin.ini capstuff section.  If "Y", it will log the comm errors.
'        Also removed auto crossover recalibration during the auto test because it was messing
'         up with the lohm correction.
'       6.70.09)
'        Added additional wait time for flow controller models (ver 7) if low flow meter
'         is under counts - it could be because there was pressure in the flow controller
'         and when valve 1 was opened this pressure went backwards through the valve and
'         caused the flow meter to read negative.
'       6.70.08)
'        Changed microflow calibration to use only 1.25 atm instead of 2 atm to make it run
'         faster.
'        Also changed parameter editor so it allows commas as well as decimal points.
'       6.70.07)
'        Fixed microflow calibration for version 7 instruments.
'       6.70.06)
'        Data averager completely rewritten. Now calculates an average CFF for the files, averages
'         dry flow data, uses the two data sets to back-calculate wet flow data, and creates a file
'         containing the averaged dry flow and calculated wet flow data.        -- MJLC 9/13/02
'       6.70.05)
'        Microflow data files now have the viscosity of the gas stored in the same
'         manner as the gas permeability data files.  This requires a change to caprep
'         if you use a gas other than one of the standards.
'       6.70.04)
'        Added function myVal(a$) which replaces the built-in val() function for cases where
'         they could use a comma in place of a decimal point.
'       6.70.03)
'        Fixed problem with long path names in test setup screen - they were not being
'         shortened properly if they didn't fit in the box because the reference label that
'         was used to determine how much space was available had been made wider.  There is
'         now a reference label that remains hidden but that is used to determine the available
'         width for any of the output boxes.
'       6.70.02) Added feature to keep appended log of the final results of all tests.
'         This is stored by user, and uses the boolean variable uselog and the
'         string variable logpath, both of which are stored with the user information.
'         The default is to not use the log.  The log path is set in a new form.  This
'         form could eventually be used for other user-specific information or
'         settings.
'        Also changed menu display of supervisor mode - there is now a "User Mode" menu
'         item that is checked if you are in user mode.  Previously, when you clicked on
'         "Supervisor Mode" it would toggle between Supervisor and User modes.  Now you have
'         to click on "User Mode" to switch to User mode.  Clicking on a mode that is already
'         selected will have no effect.  This is in response to a customer who wanted an
'         indication of "User Mode" rather than just a "Supervisor Mode" menu item without a
'         check mark next to it.  ("User Mode" now gets the check mark when you are in User mode.)
'       6.70.01) Modified font behavior of control buttons, which weren't showing characters
'         properly in Japanese. Font name can now be changed along with other text in a form, but
'         size and bold attributes remain the same. Also increased the size and spacing of many
'         elements throughout the program to correctly display foreign text strings larger than their
'         English equivalents.      -- MJLC, 7/25/02
'       6.70)
'         Started 7/8/02 - Converted to multi-lingual support by pulling text strings
'         into an external file, CapWinLanguage.ini, and including Translation.bas in
'         the project. (Note: CapWin uses TWO language files, CapWin and CapWin2.) - MJLC
'       6.69.06) Complete redo of manual control display - with all the new features
'         the old picture was getting too complicated and didn't properly display
'         some of the new configurations like topfill, liqpermonly, airtop integrity,
'         etc.
'       6.69.05) I/P regulator calibration table is now used when increasing
'         the pressure for the liquid permeability test, though it still waits
'         for the pressure to be reached so that changes in the penetrometer level
'         during the pressurization do not show up in the flow data.  (If it doesn't
'         reach the target pressure, it will increment the regulator a little more,
'         but it won't decrement the regulator if it is too high.)
'        Also, the display now shows differential pressure in liquid perm tests
'         as opposed to absolute pressures.
'       6.69.04) If penetrometer is in range before filling, it now assumes
'         that the penetrometer is at least somewhat full and within normal
'         range so it doesn't require that the penetrometer go through the
'         lower half of the range before going properly to the upper half
'         of the range.  (If the penetrometer starts out empty, it can get
'         incorrect readings that may seem like a full penetrometer while the
'         magnet makes the transition from off scale to in scale, so we make
'         sure that the magnet travels through the lower half before we accept
'         any readings in the upper half as being valid.  If the penetrometer
'         is already in the upper half, then we would have to drain out some
'         of the water first, and that is a pain to do, and can't be done
'         easily in autofill machines.
'        Also, on liquid permeameters with an autopiston (of which we have
'         only made one) the autopiston doesn't come down until the penetrometer
'         has been filled.
'       6.69.03) The capcal.d8a file is now ignored if you only have liquid
'         permeability (there are no flow meters, so you can't create a
'         capcal.d8a file, so there is no sense in trying to read in the
'         file because if it is corrupt then you have no way of fixing it).
'       6.69.02) Added support for autofill valve in version 7 porometer
'         this is feature number +512.  The topfill variable is not used in
'         version 7
'       6.69.01) Added further error trapping to serial communications routine
'         for unsupported read command.  If you send a read command that is not
'         supported by the hardware, the hardware won't send anything back and
'         the program will just loop back and try the command again.  Now, after
'         10 retries, it will give you an error message and then return with a
'         simulated reading.
'       6.69) Release version shipped 5-14-02
'       6.68.10) Added cancel button to some calls for safetykeypress
'         if the piston is about to rise at the end of the test, you can't
'         cancel it.
'        Also, in the elevated pressure liquid permeability test, if you
'         abort the test, it will force a data point (and then end the test)
'         without waiting for the timer to finish counting.
'       6.68.09) Added safetydown and safetyup features for compression
'       6.68.08) Added starting pressure to condensation test
'       6.68.07) Fixed timer problem in auto test for condensation porometry.  Also added
'         new condensation parameter "min_deltap" to control how long readings continue.
'       6.68.06) Added support for valve 20 (formerly the air top integrity valve, now
'         the air top exhaust valve) for machines with liquid permeability, air top, and
'         only one sample chamber.  Valve 20 needs to remain open most of the time and
'         only close when you are doing liquid permeability.
'        Also added support for version 7 compression regulator on analog output port 3.
'         The compression regulator is not motorized in version 7 - it is another i/p
'         converter.
'       6.68.05) Added lock for P0 value after bubble point is taken during wet curve
'         This should prevent the report program from giving a value different from the
'         testing program for the bubble point.
'       6.68.04) Added preliminary auto test for condensation test using liquid vapor system
'       6.68.03) Added support for 9th valve in liquid vapor permeameter.  You can still
'         have lvperm_numvalves set to 8 and the 9th will show up.  (For now, you can use
'         any number other than 5 and all 9 valves will show up.)
'       6.68.02) Fixed colors on analog calibration form so the text is always visible
'         even with inverted color schemes.
'        Also: Now scans com ports 1 through 9 and only puts those that are available
'         in the menu.
'       6.68.01) Allow cyclic compression to also be used with autopiston machines
'       6.68) Release version shipped 3-21-02
'       6.67.08) Added more forced regulator increments when you are below the minimum
'         pressure on the dry curve and have been opening valve 2.  This is all in an
'         attempt to speed up the dry curve below the minimum pressure.
'        Also changed final result value for bubble point so it shows 3 digits after the
'         decimal - it had been only showing 2 digits, which was ok for a 100 PSI machine,
'         but for a 500 PSI machine this wasn't enough.
'       6.67.07) Added compensation for noisy pressure gauges during the dry curve
'         when you are lower than the minimum testing pressure.  This stops the noise
'         on the pressure gauge from making the program think that the pressure is
'         rising when in fact it is just jumping around.  (If it thinks the pressure is
'         going up, it won't increase the pressure on the system and will take a long
'         time to get to the minimum pressure - and we are trying to get to the minimum
'         pressure as fast as possible.)
'       6.67.06) Added initial pressure regulator increment to lohm calibration for
'         i/p regulators that have a non-zero starting point of the pressure
'         regulator calibration table.  If the regulator has a high starting point,
'         this could cause the lohm calibration routine to give up with no pressure
'         increase.
'       6.67.05) Minimum lohm_ratio is changed from 1.1 to 1.01
'       6.67.04) Added manual control for 8-valve lvperm.  There is a new variable
'         lvperm_numvalves, which is either 5 or 8.
'         Also added lvperm manual pulsing routines when you click on the letter of the
'         valve.  You control the pulse width using a slider that uses global variable
'         lv_valve_pulse_timing.
'       6.67.03) The min_pressure use in the dry curve is now optional.  There is a check
'         box in the execute menu and a variable in the user group.
'       6.67.02) Fixed low flow controller stabilization routine when using a non-zero
'         minimum pressure for the bubble point test.
'       6.67.01) Renamed modes of testing to "QC" - replaces "Simple QC", "User" -
'         replaced old "QC", "User" or "Normal", and "Supervisor"
'       6.67) Release version shipped 2-15-02
'         Microflow calibration has been removed since a manual calibration at
'         the factory is more accurate than the auto calibration by the software
'         and the volumes shouldn't change unless the hardware is modified.
'       6.66.12) CalcGP rewritten to use the .cft data structure.  (It had been giving
'         incorrect results.)  It currently only supports darcy calculations.  Other
'         calculations will follow later.
'       6.66.11) Pressure is now displayed while pressurizing but below minimum pressure
'       6.66.10) Added extra increase to pressure regulator if we are below the minimum
'         pressure (and so are not waiting for stability) but opening of valve 2 doesn't
'         seem to be increasing the pressure any.
'       6.66.09) Fixed possible problem where pressure regulator would be incremented
'         too much during estimated bubble point setting routine
'         Also added support for minimum pressure in dry curve when going up
'       6.66.08) Added support for valve 20 for an air-top microflow.
'       6.66.07) Added in Howard's fix for the parameter saving routine
'         Also added darcy calculation to end of test report for liquid
'         permeability.
'       6.66.06) Changed display method for results at end of test.  Sometimes
'         it could choose the wrong display format resulting in strange looking
'         results.
'       6.66.05) Fixed overflow that could happen if your regulator was not
'         calibrated all the way to the maximum pressure for the test and it
'         had to extrapolate the table to get the count value for the maximum
'         pressure.  Now, the program will not try to use a count value for
'         the regulator higher than the maximum calibated count value.
'       6.66.04) Various buttons have been renamed and/or moved to make things
'         look more standard.  Forms with OK and Cancel now all respond to
'         "Enter" for OK and "Esc" for Cancel.  Other buttons all have
'         keyboard shortcuts for them.  Ampersands can now be put in group names
'         and sample IDs and other places and they will show up properly (rather
'         then showing up as underlines).  This should make companies like
'         P&G and H&V happy.
'       6.66.03) Group names are now always called "Group" - they were sometimes
'         called "User" or "User Group", which was confusing when describing
'         supervisor/user/qc modes of operation.
'       6.66.02) You can now set the pressure regulator increments less than 1
'         (but no less than 0.02) for version 7 since this parameter is
'         multiplied by 50 for an I/P converter
'       6.66.01) Fixed integrity test for version 7 instruments.
'       6.66) Release version shipped 1-8-02
'       6.65.14) Changed mapping of high flow meters 1 and 2 in version 7.  In
'         older versions, if there was only one high flow meter it was called
'         "high" and mapped to channel hflow1.  If there were two high flow meters,
'         the largest was called "extra high" and mapped to channel hflow2 and
'         the other was called "high" and mapped to channel hflow1.  Since a single
'         high flow meter is usually the same range as the "extra high" in a dual
'         high flow meter system, this naming was changed in version 7 so that a
'         single high flow meter is called "high" and mapped to the "high2" channel
'         and with a dual flow meter system the largest flow meter is called "high"
'         and is still mapped to the "high2" channel and the smaller flow meter is
'         called "medium" and is mapped to the "high1" channel.
'       6.65.13) Added cyclic compression to the execute menu for those who have
'         an autocompression feature
'       6.65.12) Fixed caption on penetrometer fill count targets to reflect
'         version 7 hardware count range.  Also changed initial setup stability
'         criteria for zero flow settings because of version 7 higher resolution.
'       6.65.11) Openv2pos now makes sure that the valve only opens
'         and Move2v2pos now makes sure that the valve only closes.  This stops
'         incorrect valve movement if the target byte is sent incorrectly
'         due to a communications problem.
'       6.65.10) Modified communications error trapping to avoid a possible
'         software hang.
'       6.65.09) Added support for time logging of data points if use_time=Y
'       6.65.08) Added manual data logging option to liquid vapor manual control
'       6.65.07) Added support for special autopiston clamping machine
'         Set "autopiston" to "Y" to enable.  Can't also have
'         autocompress turned on.  Must have special hardware or it won't
'         do anything.  Also enables door switch which won't allow piston
'         to clamp unless cover is closed to prevent pinching fingers
'         Saved testing parameters and software version number at end of test.
'       6.65.06) Fixed some timing problems with new equilibrium routine
'         introduced in 6.65.04.
'         Also, Preginc is now multiplied by 50 for version 7 because 20 was
'         not enough.
'       6.65.05) For version 7, the I/P converter is set to the maximum pressure
'         for the test at the beginning of the bubble point routine.  Previously,
'         the pressure on the regulator was kept slightly higher than the pressure
'         in the sample chamber, but this rising pressure (as the pressure in the
'         chamber rises) was causing false flow readings on the flow controller
'         and messing up the bubble point test for low flow rates (3cc/min).
'       6.65.04) Parameters are now independent of version and computer speed.
'         Slew count parameters are now multiplied by 3 for version 7 (new
'         ver1or3 variable).  Preginc is multiplied by 20 for version 7.
'         Eqiter and Aveiter are now in 0.1 seconds (though the special cases
'         of 0 and 1 are still usable).  The parameter editor screen reflects
'         these changes and allows non-integer values.
'         Also, fixed problem with initializing the system with the i/p converter
'         for the second half of a DUWU or WUDU test - valve 11 may have been
'         opened before the pressure in the system had bled off enough so the
'         fail-safe procedure in the instrument would have closed the valve
'         again.  The valve 11 opening has been moved to after valve 2 is closed
'         and valve 11 is again set in the "load sample" routine just to make
'         sure.
'         Also, a 3 second delay at the end of the "initialize system" routine
'         was removed because it probably did nothing but make test setup
'         take longer.  It was originally put in the QBasic code so that a
'         display message would stay on the screen long enough to read it.
'       6.65.03) Changed version 7 protocol so return values are now 6 bits.
'         This makes the return values all printable characters and eliminates
'         the problem with null return values messing up the sync.  This was
'         originally implemented in l3perm.  This means that return values that
'         were already 1 byte printable character are unchanged but that return
'         values that are 1 byte binary are now returned as two characters and
'         normal analog readings that are 2 bytes binary are returned as three
'         characters.  There are two new version dependent global variables to
'         make this a little easier.  Ver2or3 is set to 2 for version 6 and 3 for
'         version 7.  Ver1or2 is set to 1 for version 6 and 2 for version 7.
'         The debugcom routines do their own binary input and do not use this
'         method, but since the goal is to minimize com errors when using the
'         normal testing routines and to maximize the error detection when using
'         the com testing routines, this is fine.
'       6.65.02) Added system_font variable and selectfont form to allow user to
'         set the font to use for all forms.  This should allow users in countries
'         that use non-ascii characters to properly enter their own values.
'        Also added support for saving a named .cal file for the lohm table
'         and re-loading the named file to overwrite the lohmtable.cal file.
'         Named lohm table .cal files are stored in the parms directory so we
'         don't have to make a new subdirectory.
'        Also changed all GPPS calls that return strings to use a null trim
'         function and not use the returned length of the string.  GPPS returns
'         an incorrect length (too long) if the return value contains multi-
'         byte characters (such as Japanese or Chinese), and this padded the
'         string with extra null characters which messed up string handling
'         functions.  This showed up in path names in the file selector.
'       6.65.01) "Please Wait" form now doesn't show if you are just saving your
'         test setup and not actually running a test.
'       6.65) Final version shipped on 10-01-01
'       6.64.21) Added simple QC mode - displays user groups and automatically
'         runs the default test in that group with no prompts.
'       6.64.20) Fixed problem with first data point of gas permeametry being
'         overcorrected by a high lohm value - would most likely happen with
'         surface area testing.
'       6.64.19) Manual control of liquid vapor now correctly closes valve E
'         when it starts up.  (Otherwise, if you left valve E open and closed
'         from manual control, when you re-entered the screen would show the
'         valve closed when it was in fact open.  If you then clicked on it,
'         it would either close and the display would continue to show it closed
'         or it would stay open and the display would then show it open.
'       6.64.18) Fixed problem introduced in version 6.64.16 that would cause
'         a subscript error on all tests that did not use the temperature
'         probe (even if there was no temperature probe in the instrument).
'         The feature to remember the testing temperature was also triggering
'         the use_temperature variable incorrectly so it always thought that
'         there was a temperature probe being used.
'       6.64.17) microflow porometry now doesn't need the pressure gauge to be set
'         with a zero value of 0 - you can set the zero value to 14.7 so that it
'         reads absolute pressure correctly and it will still work.
'       6.64.16) First letter of alternative fluid name was being messed up - this
'         is now fixed.
'        Also, testing temperature for liquid permeability is now remembered (for
'         those few machines that have this feature).
'        Also, liquid permeability default values now show up without the leading
'         space and are fully selected to make it easier to change them.
'       6.64.15) Allow autofill liquid permeability test to take up to 5 minutes for
'         the initial fill before giving up (autofill has to fill the entire sample
'         chamber, so it can take longer for the penetrometer to get in range).
'        Also, the prompt at the end of the autofill has been changed to make it
'         clearer.
'       6.64.14) Fixed manual control for topfill autofill machines - it was still
'         indexing the drain valve as valve 12 when it was in fact valve 3.
'        Also fixed manual control when temperature controller was turned on -
'         the temperature text box was sending keypresses to the form, but the form
'         already had seen them because of the preview feature turned on, so some
'         keypresses were being processed twice whenever the cursor was in the
'         temperature setting text box.
'       6.64.13) Machines with solenoid valve 2 (WESA) now can run a lohm calibration.
'       6.64.12) Added support for new valve E in manual control of liquid vapor
'       6.64.11) Apply 1/20 span offset to open limit so that we don't have the
'         problem in manual control of having the valve stop at 99.5% open.  This
'         offset was used in the older version 5 machines, but never added to the
'         version 6 machines until now because you could always change the limit in
'         the ini file, but now the program automatically re-reads the open limit
'         (and the close limit) every time the instrument is initialized.
'       6.64.10) User settings now remember if you want to show results at the end
'         of the test.
'       6.64.08) If your output data file has the name "do not save data file.cft"
'         then the data file will not be saved.  (If there actually exists a file
'         with this name, you will not be asked if you want to overwrite it unless
'         you select it again in the file selector since the file selector has a
'         a built-in prompt for overwriting existing files, but the capwin program
'         will not add an extra prompt when you start the test.  The existing file
'         will not be modified.)  The lastdata.cft file will still be written with
'         the data from the last test run.
'       6.64.07) Fixed problem where gas permeability test wasn't saving the
'         information about which gas was selected - it would always default to air.
'       6.64.06) Improved low flow controller initialization so it works better at
'         low flow rates (such as 2 cc/min).  It was giving false bubble points.
'       6.64.05) Maximum count and interation values are now higher in version 7
'       6.64.04) Dry down now works better in version 7.  The regulator will
'         be decreased steadily, and at the same time valve 2 will be closed
'         in amounts that will bring it to fully closed by the time the regulator
'         reaches 0.
'        Also: Fixed problem with pressure rising too much when the low pressure
'         gauge goes off scale so fast that there is no time to do the automatic
'         crossover calculation - if this happened during normal equilibrium then
'         the program would keep incrementing the pressure (opening valve 2 and/or
'         increasing the pressure regulator) until the low range of the high pressure
'         gauge reached the same count value as the previous good reading of the
'         count value of the high range of the low pressure gauge.
'       6.64.01) Added communications debugging commands for serial port on hardware
'         version 7 machines.  This should give us information on the latency of the
'         serial port and be able to test transmit and receive independently and give
'         reliable statistics on how the serial port is working.
'       6.64) released 7-5-01
'       6.63.10) airtop integrity now works with integrity porometry testing.  If
'         the maximum flow rate parameter is set at or below the range of the
'         integrity flow meter, then variable integrity_porometry is set to true
'         and during the porometry (and gas permeability) test instead of reading
'         the high flow meter it will read the integrity flow meter.
'       6.63.9) Added support for valve 20 - this is for new airtop systems with
'         an integrity flow meter.  Valve 20 stays open at the bottom of the air
'         sample chamber unless you want to divert the flow through the integrity
'         flow meter, and then valve 20 closes so all the flow has to go back
'         up inside the cabinet to the integrity flow meter.  This is activated
'         when airtop and integrity are both true.
'        Also - when you click on the integrity flow meter on an airtop machine,
'         if valve 20 is already closed, it will open valve 20 and stop reading
'         then itegrity flow meter.  (Previously, you had to click on the low flow
'         meter.)
'        Also - valve 2 limits are now automatically updated every time the machine
'         initializes.  You no longer need to do the calibration manually, though
'         you still can (it will let you know that you don't need to do this.)
'         This only applies to version 6 and higher machines.  For version 5.31 and
'         lower you still need to do the valve limit calibration whenever you change
'         valve 2 (and the calibration message doesn't tell you that you don't have
'         to do the calibration).
'       6.63.8) Chamber selection can now be done in user mode - it was previously
'         restricted to supervisor mode.  The chamber selection is now done using
'         a new form, rather than being in the menu, as this makes it much easier
'         to set chambers.
'        Also: New groups will inherit default parameters from current group.  This
'         wouldn't work if you selected a different group and then created a new group
'         without running a test first - the selection of a group was not loading in
'         the default parameters until the next test was run, so when the new group
'         was created it would get its parameters from either the group that was current
'         when the program was started, or the parameters from the last test that had
'         had been run.
'       6.63.7) Gauge range changes in manual control are now delayed until after
'         the current reading is finished.  This fixes a slight problem where
'         the next reading would be messed up when you clicked to change the range
'         of a gauge.  The count value would be for the previous range but the
'         calculated value would use that old count value with the new range.
'       6.63.6) I/P converter now has pressure regulator calibration enabled.
'       6.63.5) Fixed problem where if you have only one high flow meter it would be
'         possible to get an error message saying that hflow% was set to 2 (it should
'         only be set to 0 or 1 for one high flow meter) at the beginning of the test
'         if you run the test after running a lohm calibration.  If you restart the
'         program after running the lohm calibration you will not see this problem.
'         The Lohm calibration function was leaving a variable set incorrectly, and
'         the load sample function was assuming that the variable was set properly.
'        Also set the lohm value when the lohm table is zero length to 0.  This will
'         solve the problem when there is a large sample chamber with a very low
'         pressure gauge mounted in the chamber (for gas permeability tests) and there
'         is no pressure drop measurable during the lohm test.
'       6.63.4) When you select advanced settings for one test it would now reverts
'         back to the simplified settings for the start of the next test.  Before,
'         it would sometimes stay on advanced settings.  This broke in 6.62.06.
'         This problem was caused by a variable local to the form retaining its value
'         even after the form was unloaded and then loaded in again.  The program
'         assumed that the variable would be cleared by the unload process and when it
'         wasn't it assumed that the form had already been loaded and didn't bother to
'         reset the advanced view to the default value.  There is now a global variable
'         that is set by the routine that calls the testing screen to let it know that
'         this is the first call for a given test.
'       6.63.3) Preliminary support finished for version 7 hardware - auto testing
'         and calibration should all work now.
'        Also: low flow calibration is stored in capcal.tmp while it is being done
'         and it is only copied over to capcal.d8a when it finished correctly.
'       6.63.2) CAPWIN.INI new option: preloaded_sample=Y means that we can assume
'         that the sample is already loaded and don't have to ask for anything to
'         start the test after the press "Start" on the test setup menu.
'       6.63.1) Valve testing ("t" in manual control) will now only work if you
'         enter a positive number for the number of times to test valve 2.  It
'         used to have a default value of 100, which would be used if you cancelled
'         the number input box, and would also allow 0 for a valid number of tests.
'        Also: Pressure unit menu now has a check mark next to the current unit.
'        Also: Tab orders fixes for some forms to make it easier to use keyboard
'         to set up testing.  Some more alt-key shortcuts added as well.
'        Also: Now remembers last settings for leak test, curve fitting, and elevated
'         pressure liquid permeability.  These are remembered independently for each
'         user.
'        Also: Data editor now knows if the current file has been modified and warns
'         you when you try to exit the form or load a new file if you haven't saved
'         the changed data.
'        Also: MFP is now reported at end of porometry test (if they want data)
'       6.63) Released 6-12-01
'       6.62.07) Fixed problem where dry curve sample load would always prompt
'         for loading even in multiple runs.  This affected dry up/wet up, and
'         gas permeability.
'       6.62.06) Test setup screen now properly skips unselected chambers.  It had
'         been showing setup screen for all chambers between the lowest selected
'         and the highest selected, but it will now skip unselected chambers in
'         the middle (such as 1, 3, 4).
'       6.62.05) User is now always prompted to put in saturated sample in the
'         middle of a dry up/wet up test with multiple chambers, but is not
'         prompted to put a dry sample in the next chamber because that is
'         assumed to have already happened.
'       6.62.04) Fixed problem where bubble point would sometimes be skipped in
'         multiple chamber testing.
'       6.62.03) Load chamber message box now tells you which chambers you are using.
'         and if you abort a test you will then get the low sample message for the
'         next chamber.  (If you don't abort, the next chamber will start automatically.)
'       6.62.02) Fixed problem where multi-unit tests were pausing to ask user to
'         install sample in next chamber - user should only be asked at beginning
'         to install samples in all chambers, and should only be asked at the very
'         end to remove all samples (except for cases where there is large amounts
'         of sample preparation, such as for liquid permeability testing).
'       6.62.01) Fixed display problem in multi-unit instruments when the first
'         unit is not selected - the test setup screen was not being initialized
'         properly.
'       Also fixed problem where bubble point would be skipped on subsequent units
'       6.62) Release version 5-31-01
'       6.61.11) Data editor now correctly saves files with moved dry curves.
'       6.61.10) Lohm table now has a third column for the low pressure gauge
'         if there is a low pressure gauge.
'       6.61.09) Testing of parameter files now doesn't give false error if a parameter
'         file is invalid but the type of test being performed doesn't use this parameter
'         file.
'       6.61.08) Fixed multi-chamber lohm calibration - may also have affected running
'         the lohm calibration more than once without shutting down the program.
'       6.61.07) When starting a test, the program now makes sure that the parameter
'         files exist - if they don't (because of a network problem or removed directory
'         or some other reason) then you will be warned (as opposed to having the program
'         crash).  If the parameter file is no longer available when the program needs to
'         open it, there will also be a warning message.
'       Also, at the end of the test, if it can't copy the temporary data file over to
'         where you said you wanted it, it will not crash but will ask you where to put
'         the file and try again.
'       Also: the "autocrossover.txt" file will only be created if the variable
'         "crossoverdebug=Y" is in the capwin.ini file.
'       6.61.06) Added 2 second delay when closing valve 12 on a lowered penetrometer
'         system during an elevated pressure test.  This may help with air getting
'         below the sample for this unusual "water from the bottom" test.
'       Also: Temporary creation of autocrossover.txt file when pressure gauge
'         range changes and crossover is recalculated.
'       6.61.05) Speed up old hardware analog converter (version 5.31) - it was delaying
'         longer than it needed to to wait for analog signals to settle.
'       Also: If there is a path name in the command line to the executable, this will
'         be used as the path to the capwin.ini file.  If there is no path name, or there
'         is no capwin.ini file at the path name, then the normal procedure will be used.
'         This makes it easier to run the interpreted version from any path.
'       Also: On multi-chamber systems with compression, the compression feature
'         is only used for chamber 1 - other chambers are normal.  This may have to
'         change if we ever make a multi-chamber system where there is compression
'         on more than just chamber 1.
'       Also: On multi-chamber systems, the lohm calibration will only use those chambers
'         that are currently selected, and generate a lohmtable.cal file for each
'         chamber (with a number inserted before the ".cal" for chambers 2+).
'         When a test is run, the proper lohm table will be loaded for each chamber.
'         If a table is not found for the current chamber, the program will try to use
'         the one for the previous chamber and keep going back until it reaches chamber
'         1, which uses the default lohm table which must be there.
'       6.61.04) Fixed overflow problem if you used a sample ID (or other text input line)
'       that could be evaluated to a very large number (like 3e456).
'       6.61.03) Added support for 5 chamber sequential instrument
'       6.61.02) Fixed subscript out of range error when logging lohm values
'       6.61.01) New variable max_liq_pres stores the maximum liquid permeability pressure
'       default is 200 PSI (which was previously hard coded).
'   6.61 released 4-29-01
'       6.60.19) Updated some things from Jeff - you can now abort the lohm calibration
'       Release versions will now only use the first two version places
'       The third version place will be 0 on released versions.  Modifications
'       to released versions will use the third place, and when the modifications
'       are complete the third place will be reset to 0 and the second place will
'       be incremented by one.
'       6.60.21 fixed parameter file printout problem - also moved parameters around
'           so that maximum pressure is now in a special part since it is no longer
'           used for testing and is only used for lohm calibration when there are two
'           regulators.
'   6.60 1-24-01 first version recompiled in vb6
'       1) By using VB6 and native code compile, the program runs faster
'       2) Changed from CV to Lohm calculations - improves correction of
'           high flow/low pressure samples
'       3) Added "Prefilled" button to penetrometer fill routine that bypasses
'           the check for initial fill.  This is needed when the penetrometer is
'           already filled before you start the test.  The initial fill check
'           requires that the penetrometer count value is greater than 10,000
'           before it starts reading anything.  This was put in to fix the problem
'           with false readings of full when the magnet is just entering the range
'           of the penetrometer.
'       4) In elevated pressure liquid permeability, we won't increment the regulator
'           for the first two data point faster than once every 1 second.  This should
'           stop the pressure from rising too fast.
'       5) You can now abort a microflow test while the pressure is initializing
'       6) The pressure target is now displayed during elevated pressure liquid
'           permeability while the regulator is being incremented.  This is for
'           debugging purposes and may be removed later.
'       7) Fixed problem where elevated pressure liquid permeability test would
'           record an inaccurately high atmospheric pressure before you actually
'           start the test and then because of this not take any data at really
'           low pressures at the beginning of the test.
'       8) If they have a bottom fill penetrometer (the pen20500 value is negative)
'           then the drain valve (12) is not really a drain valve but is a chamber
'           isolation valve between the chamber and the penetrometer.  In this case,
'           we want to leave this valve closed at the end of the test so we don't drain
'           out all the liquid from the chamber so they can run another test easily.
'           For now, this valve will have to be opened in manual control if they really
'           want to drain things out.
'       9) All output data files are now initially written to the file LASTDATA.CFT
'           in the same directory as the executable program.  When the test is finished,
'           this file is copied over to the actual output file that the user requested.
'           The LASTDATA.CFT file remains so if they can't find their output file they will
'           at least be able to recover the last test data.  This is also needed for high
'           security situations where the raw data file is written to a non-modifiable
'           disk partition so it can't be modified after the run is over.
'       10) For bottom fill penetrometer, at the end of an elevated pressure test it will
'           now close valve 12 before venting the penetrometer and also wait 2 seconds
'           for the valve (usually an air operated ball valve) to fully close.
'       11) In preparation for version 7 hardware support, all references to absolute
'           count values have been removed.  The variables DAC_x are now set based on the
'           hardware version.
'           DAC_under = the minimum count value possible (1 for version 6 and 0 for version 7).
'           DAC_over = the maximum count value possible (23000 for version 6, 65535 for version 7).
'           DAC_zero = the normal count value for zero volts (500 for version 6, 2000 for version 7).
'           DAC_two = the normal count value for 2 volts (20500 for version 6, 62000 for version 7).
'           DAC_span = DAC_two-DAC_zero (20000 for ver. 6, 60000 for ver. 7)
'           All references to the feature number should be inside the routine that reads the capwin.ini
'           file.  This should set all booleans based on the feature number.  The bits in the feature
'           number has changed slightly between versions 6 and 7, so haveing the decoding take place
'           in one place will make things easier for the future.  The "feature" variable is now local
'           to this procedure, while a new variable "capwin_feature_number" is used to store
'           the feature number that is read in so it can be compared with the actual hardware feature
'           number when communications is established with the instrument.  Do not use this new global
'           variable in determining which hardware features are available anywhere outside the capwin.ini
'           reading routine.
'           While doing this, variables that should have been boolean but were integer because they
'           were introduced in an older version of the language that did not have boolean variables
'           have now been converted to booleans.  String variables that only held two possible values
'           (such as "Y" and "N") have also been converted.  Strings that had three possibile values,
'           such as "Y" for yes, "N" for no, and "E" for exclusive, have been converted into two
'           booleans - one for enable and one for exclusive.
'       12) Fixed bug where program would crash with a type mismatch if you clicked on the "Edit Parameter
'           File" button and the current focus was not on either the wet or dry parameter file and both
'           the wet and dry parameter files were visible and they pointed to different parameter files.
'       13) Converted all variables that can hold count values from integer to long - this is so they will
'           work with the version 7 hardware that returns count values between 0 and 65535.
'       14) Feature number on version 7 hardware always has the highest bit set.  Version 6 always has
'           the highest bit clear.  Some version 6 hardware (old) may not have a feature number, and this
'           has been allowed, but all version 7 hardware must have a feature number or it will not work.
'           This will alert you if you connect the wrong machine to the serial port.
'       15) The CalibBoard form no longer uses a timer - it now reads as fast as it can.  It also has the
'           capability of reading only one of the two calibration values (+2 and gnd) so you can make sure
'           the reading is stable.  On version 7 you can also set the delay and averaging settings for the
'           most steady signal.  These delay and averaging settings are now stored in the capwin.ini and
'           updated when the instrument is initialized.
'       16) Added statistics for how fast the analog readings are coming in to the
'           timer in the manual control window
'       17) Removed "ga" protocol for version 7 to solve serial port latency problem.  Modern serial ports
'           do not need this protocol, which was implemented to get around single byte buffers in older pc
'           serial ports.  Now that version 7 hardware is so much faster than version 6 that the latency of
'           the serial port is making a big difference.  (It makes a small difference in version 6.)
'       18) New boolean low_flow_controller if the low flow meter is actually a controller.  This is default
'           in hardware version 7 and never present in hardware version 6.
'   6.54.48 work in progress started 11-20-00
'       1) Moved release of compression pressure at the end of the test until
'           after the regulator is zeroed and vent valve is opened.
'       2) Valve 1 is now opened during the excersize valve 2 routine.  This
'           will help vent any high pressure gas between valve 1 and the needle
'           valve that would otherwise have to go through the needle valve
'           and low flow meter, which would cause the low flow meter to take
'           longer to stabilize during the initial zeroing.
'       3) Estimated bubble point can now be done with elevated pressure bubble
'           flow.
'       4) The bubble point test is timed from when the sample is first loaded to
'           when the bubble point is recorded.  This time interval is reported along
'           with the bubble point on the status line.
'       5) Added leak test to liquid vapor perm evacuation routine - pressure must
'           not go up by more than 0.1 torr in 10 seconds or the test will not start.
'       6) Modified caption of new user units conversion factor entry form to make
'           it clearer which conversion factor you should use (which is the amount in
'           your new unit that is equal to 1 PSI).
'       7) The curve fit routine now lets files have flow values that
'           are going the wrong way.
'   6.54.47 11-17-00
'       1) Added support for experimental liquid vapor permeability test
'           using two auxillary pressure gauges and four valves
'           This is enabled by the lvperm_enable boolean set in the ini file
'           The two pressure gauges are calibrated with aux_p1_span and
'           aux_p2_span, which both default to 100 torr.
'           The lvperm_enable is set to "E" in the ini file, then it is
'           an exclusive lvperm with no other capflow features
'       2) The debug button for storing raw data with a pressure increment is
'           now turned on by the debug_button_enable boolean in the ini file.
'       3) Added support for liquid permeameter with hydrohead - looks like a
'           bubble point tester (with low flow meter and solenoid valve 2) but
'           also has a penetrometer.  Does not have valve 4 since there is no
'           way to send the pressure to the bottom of the sample chamber.  This
'           type of machine can not run bubble point tests, but can run hydro-
'           head tests.
'       4) Added support for I/P converter controlled regulator - feature 2048
'           boolean variable ip_reg_enable - works with cv calibration, flow
'           flow calibration, and gas permeability.  Other functions not tested
'           yet.
'       5) Improved transition between bubble point and wet curve - on some systems
'           it would hang up while lowering the pressure regulator
'   6.54.46 10-6-00
'       1) Modified release of sample for high pressure liquid permeability
'           test when using an automated compression system - it now releases
'           the air pressure on the liquid first and then lifts up the compression
'           piston - it was causing leaks around the sample chamber at the end of
'           the test for high pressure testing.
'       2) Fixed problem in data editor when you tried to import a dry or wet
'           curve with fewer data points than the current curve it would mess up.
'   6.54.45 Released 10-2-00
'       1) Fixed bug in curve fit routine where if the dry curve (coming down)
'           stopped too early it could eliminate some of the wet curve data points.
'       2) When running dry up / wet up porometry test, the bubble point pressure
'           would incorrectly be reported with the label "Hydrohead Pressure"
'           on the testing screen - this is now fixed.  (The data would be correct,
'           only the label would be wrong.)
'       3) Added supervisor variable - when true ("Y" in INI file, which is default -
'           everything is normal.  When false ("N"), all complicated things are hidden
'           so you can only do the default test in any of the groups.
'       4) When you close the test setup screen (with the "CLOSE" button, not with
'           the "X" box) you are now allowed to save the changes you may have made to
'           the test setup.  This allows you to pre-set things without actually running
'           a test.
'       5) Test Setup Screen changed to allow selection of which items are shown
'           in QC mode, no more double-clicking (single click will do), takes up less
'           room (won't go off screen on small monitors), all user information now moved
'           to a common routine, etc..
'       6) Moved subroutines that are common to both the control and report programs
'           into a common module.  This will make updating both programs easier.
'       7) Data File Editor updated so it will work with all data file types.  It now
'           uses the common files and cfttype data structure.
'       8) Pressure hold test now uses the dry parameters line in the test setup
'           screen to hold the pressure and testing time, plus the new parameters for
'           initial delay time and maximum rate of pressure drop (for pass/fail report)
'       9) Added button to manual control for debugging purposes.  This is normally hidden,
'           but when it is turned on it will activate a debug program that can be used
'           by our engineers to do special projects.  Initially, this routine will collect
'           data on pressure, flow, and time as fast as it can after incrementing the
'           pressure regulator.  This data can be used to analyze how the pressure and flow
'           reach stability.
'   6.54.44 7-17-00
'       1) Added reg_zero_time to capwin.ini - defaults to 8 seconds.
'       2) End of external Hydrohead test can now be triggered with space bar
'       3) CV Interpolation routine fixes - it could cause problems if the CV value goes
'           down when the flow goes up.
'       4) Fixed problem in bubble point testers (with no high flow meter) where
'           startf parameter was messing up in the parameter editor.
'       5) Now works with 10 chamber bubble point tester (CHAMBERS=10)
'       6) Fixed problem where if you closed the test selection screen using
'           the close button the test would start - now it acts just as if you
'           clicked on the cancel button.
'       7) Microflow test now waits for differential pressure gauge to return to
'           initial value after closing of venting valve before it starts the timer
'           to measure the flow rate based on the change in pressure.  This should
'           fix the problem with a very small pressure gauge going slightly negative
'           when the venting valve is closed.
'       8) Pressure gauge calibration for a single gauge less than 150 PSI now
'           works better - it would previously get the cross-over pressure correctly
'           but could over-shoot when doing the secondary check due to too much
'           pressure (it incremented the regulator too much).
'       9) Hydrohead now works with airtop porometers that don't have liquid
'           permeability (as long as proper ROM is used with Hydrohead support)
'           If your machine has air from the top for the normal air type tests,
'           it will now use this air supply to run the hydrohead rather than sending
'           all the air pressure through the penetrometer.
'   6.54.43 4-19-00
'       1) Changed EFD to jump to near but less than maxflow before taking data.
'       2) Fixed problem with curve fit program where sometimes it would give
'           a subscript out of range error.  This error was introduced in
'            version 6.54.42 when the random access files were removed.
'       3) Fixed problem with another "subscript out of range" error when you
'           would abort a two-pass test during the first pass.
'       4) The sqrt type of "linear dry" is now working.
'       5) Rewrote valve position communication routine (and temperature
'           setting routine) that used binary 16-bit values encoded into
'           string variables.  Some values would not work in these routines
'           when running an operating system that used 16-bit and 8-bit
'           characters (such as Chinese, Japanese, Korean, etc.).  The RS232 routine
'           now has a second version that takes in a 16-bit number as a
'           new final argument.  The string is sent normally, followed by
'           the 16-bit number.  In the manual control tempset_click routine,
'           the target temperature is now encoded into the pending string
'           as base 128 instead of base 256 to avoid this error as well.  Since
'           the maximum value is 9999, this still fits into 2 characters, but
'           neither of the characters will ever go above 127.  This hasn't been
'           tested as we don't have a temperature controller enabled porometer
'           in Asia.
'       6) Modified high flow zero correction slightly in pressure gauge
'           calibration function to make it more consistant.  Probably won't
'           affect anything on how the calibration works.
'       7) Changed manual control display format slightly so that pressure units
'           greater than 4 characters (such as kg/cm2) will show up properly.
'       8) Fixed leak test so it properly uses non-PSI units if selected.
'       9) Fixed problem where "Execute" menu would get re-enabled during test
'           setup.  This could allow the user to start another test at the same
'           time.
'       10) The cancel button on the test selection screen now responds to
'           the "Esc" key, as all "Cancel" buttons should.
'       11) Fixed problem with implementation of autofill and airtop - if you
'           have feature 16 (meaning airtop) and not autofill you would get an
'           incorrect error message
'       12) Changed pressure drop test control slightly to improve performance.
'   6.54.42 2-7-00
'       1) Fixed external frazier pressure gauge implementation
'       2) Added SC diameter to output. Changed Pressure drop output.
'       3) Removed all random access files and replaced them with
'           dynamic arrays.
'       4) Added third option (sqrt-based dry curve) to the linear
'           dry option.
'       5) Added support for second penetrometer, which adds valve
'           19 (second penetrometer venting valve) and penetrometer
'           readings on the auxillary analog port.  To enable this,
'           you need to have the auxillary port turned on (feature
'           64) and have "Second_Penetrometer=Y" in the capstuff
'           section of capwin.ini.  You also need entries for
'           2PEN500, 2PEN20500, 2CSECAREA, and 2PSIPERCM for the
'           second penetrometer.  If the 20500 count value (PEN20500
'           of 2PEN20500) is negative, this means that the penetrometer
'           is below the level of the sample chamber and you can't do
'           a normal ambient test - if you select an ambient test, the
'           program will use the bublflow value to flow gas slowly into
'           the top of the penetrometer.  New boolean variable called
'           special_ambient is true when using a penetrometer that can
'           do this special ambient test.  If you are doing such a test,
'           the wet parameter file is selectable from the setup screen,
'           and the penetrometer fill procedure has the intial fill
'           portion bypassed since the initial fill will happen before
'           you seal the sample chamber.
'       6) Added support back in for autofill.  If "Autofill=Y" is set
'           in the INI file, then the 16 of the feature number will be
'           used for autofill and not for airtop.  This is needed so
'           that older autofill machines (we only made one or two) will
'           still work with the new software.
'   6.54.41 12-20-99
'       1) Modified curve fit routine so that first and last points
'           in each curve do not get modified
'       2) Fixed interpolation routine (in curve fit and data average)
'           so that it will correctly work with dry curves with only
'           two data points.  It was dropping the first (highest
'           pressure) data point and getting an incorrect last data
'           point.  (This last data point was then truncated by the
'           report program since it was in the wrong direction, so we
'           didn't notice it.)  This fixes a rare problem with lockup
'           of the curve fitting routine.
'       3) Pressure gauge calibration now properly compensates for
'           the initial zero flow rate of the high flow meter when
'           calibrating the crossover point of the high pressure gauge.
'       4) Modified external hydrohead test so that it vents the sample
'           at the end of the test.
'       5) Fixed problem where selecting user-defined units would cause
'           the flow rate to be invisible during the next test run.  The
'           user defined unit would end with a null character which would
'           cause the test status screen to stop displaying information
'           after the pressure value.
'       6) Added new capwin.ini value "AirTop", which should be equal to
'           "Y" if the air comes from the top of the instrument.  This
'           changes the display in manual control.
'       7) Changed use of the +16 value for the feature number.  It used
'           to mean that autofill was present.  We haven't made an auto-
'           fill machine since version 5 hardware, and that didn't use
'           the feature number, so we have never used the +16 feature
'           value.  A +16 now means that there is a liquid drain valve
'           on valve 12.  This is only possible on an AirTop machine
'           since that would have drain valve 3 above the chamber but
'           on a penetrometer system there would still have to be a
'           drain valve (12) on the bottom of the chamber.  This valve
'           (12) must remain open at all times except for when you are
'           filling or refilling the penetrometer.  If +16 is present
'           and +8 is not, then we are still a penetrometer system, but
'           the drain valve (12) is added and it is a solenoid valve.
'           If there is both a +8 and +16, then the drain valve is
'           motorized and the position and limits of the drain valve
'           are read using the analog input normally used for motorized
'           valve 10.  (This means you can't have both a motorized valve
'           10 and a motorized valve 12.)  If we ever have to have both
'           types in the same machine, we will have to do something
'           different.  Since we will be switching to hardware version
'           7 soon, and everything will get re-mapped, this shouldn't
'           be a problem.
'   6.54.40 11-18-99 to 11-23-99
'       1) Added in pressure drop test
'   6.54.39 9-8-99 to 11-18-99
'       1) Changed hydrohead display of inH2O to cmH2O
'       2) Increased initial regulator clicks for high flow bubble point
'       3) Fixed cv calibration and correction code
'       4) Fixed starting X-Axis label when non-PSI units
'       5) After bubble point, regulator is reduced until low flow is less
'           than 1 cc/min or the pressure starts falling, so that when valve
'           2 starts opening the pressure doesn't build up too fast.  This
'           gives more data points for some samples
'       6) On start of up curve, don't increment regulator more until valve
'           2 is opened at least 200 counts from the starting position.  This
'           prevents the pressure from building up too much before the valve
'           has had a chance to really open and get some flow through it.
'       7) When opening valve 2 during test, each interation must open it at
'           least one count, even if the last pulse overshot the target.
'       8) Wet Up/Linear Dry test now has option to use permeability curve
'           for dry in place of linear curve.  It also plots the dry curve
'           on the screen.
'       9) Bubble point is now correctly taken at first pressure where F/PT
'           went past the target and stayed there - there were some cases
'           where the bubble point pressure would be slightly higher if
'           the pressure kept going up after reaching the F/PT target but
'           before it had been confirmed.
'       10) Finished correcting continuing QC WESA error.
'       11) Added in ability to report permeability/specific surface area/
'           bubble point at end of test. Uses the DFG parameter in caprep.ini
'           to determine what kind of permeability to show.
'       12) Fixed calibration problem when using solenoid valve 2
'       13) If using integrity and compression at the same time, integrity
'           flow meter is read using the low range of the low flow meter
'       14) Fixed "stop at bubble point" in wet up/dry down test - it wasn't stopping.
'       15) If auto-compression pressure is set to 0, the piston will remain
'           retracted and the compression regulator will be set to 20 PSI to keep
'           the piston up.  Previously, the piston would be left up but the pressure
'           would be set to 0 and on some large pistons this would cause the piston
'           to creep down due to gravity.
'       16) InputMode of MSComm set from text to binary.  This fixes a problem with
'           delays in Japanese Windows (which can have two-byte characters, and if the
'           binary data we want to get back happens to look like the start of a two-
'           byte character, then the control tries to wait for the second byte and
'           waits for the timeout value which delays the entire program.  Also switched
'           to sending a byte array for output so that it works with non-ascii characters
'           (specifically, so the valve position command "G" will work since it sends a
'           16-bit binary value for the target position, and if one of these bytes happened
'           to be a start of a two-byte character sequence, it wouldn't work properly
'           in Japan.
'       17) Added support for older IBM Interface Board instruments - version 5.3 and 5.31
'           which requires new 32-bit DLL for handling INP and OUTP commands to the board.
'           This is not working fully yet.
'       18) Added the "T" command to determine test enable properties of the ROM.  If the
'           instrument is old, the test enable value will default to all on.

'       The following feature is currently disabled:
'       x) Added experimental display of three different dry curve functions
'           while test is running - F/(P-P0), CV, and Darcy-like F/((P/P0)^2-1)
'
'   6.54.38 7-7-99 to 8-19-99
'       1) Added in input for mass and density for surface area analysis
'           tests using gas permeability. This corresponds to a change in
'           caprep that allows the data to be read in for doing the surface
'           area analysis.
'       2) Added in WESA global to indicate if the machine is only a WESA (E),
'           can do WESA (Y), or cannot do WESA (N).
'       3) Added in machine serial number to data file.
'       4) Added support for feature bit 1024 - dual regulator.
'           This adds valve 17, which switches between the two
'           regulators.  Also adds "reg1pmax" for maximum pressure
'           on regulator 1.  Also adds valve 18, which is parallel
'           to valve 1 but is used with you are using regulator 2
'       5) Changed CV calibration so that it uses both high flow meters
'           (if present).  This should help with low flow CV values.
'       6) CV calibration and flow calibration now give status messages
'           on the screen while they are running.
'       7) Added experimental external hydrohead chamber - remaps valves
'           12, 13, and 14 for pressurize, fill, and vent functions.
'           Right now the fill function is not implemented - filling
'           is done manually.  Note that this won't work with a microflow
'           machine unless valve 14 is used to vent both at the same time
'           (which could work since no one would use both fixtures at the
'           same time).  INI file has "External Hydrohead=Y" to enable
'           this feature.  (You also need to have the proper ROM and
'           external hardware or this won't work properly.)
'       8) Added in Quality Control WESA code which outputs surface area
'           and average particle size when the test is done.
'       9) Corrected QC WESA code.
'   6.54.37 6-18-99 never shipped, changed to 38 when WESA added
'       1) On systems with motorized regulator, cleanout and cv test
'           will now start with regulator open to SHFP point (where
'           calibrated low flow rate is above 0.2 cc/min).  This should
'           skip over the zero point of the regulator.
'       2) Fixed alternative fluid output to data file - internally
'           the viscosity is embedded with the fluid name.  This is
'           splint into two lines for the output data file.  Curve
'           fit, editor, and averager all thought that the data file
'           had the embedded viscosity, so they wouldn't work
'           properly.  The normal test data file output was fine.
'   6.54.36 5-12-99 to 6-16-99
'       1) Fixed display problem in data editor where dry curve would
'           be stored incorrectly
'       2) If last file name ends in at least three digits, it will auto-
'           increment the number for the next test.  If necessary, it will
'           add another character to the file name so x999 -> x1000.
'           This can be turned off from execute menu for each user.
'       3) Display of file names in the test setup form now will use a
'           shortened version of the file name if necessary so that the
'           text does not wrap to the next line.
'       4) Auto Increment and Auto Advanced settings are now copied to
'           a newly created user group (along with other settings).
'       5) Test setup screen now correctly shows wet up/linear dry test
'           when you change the group to one that has that as the last test.
'       6) Selection of users is now sorted and scrolls horizontally.
'       7) User group names can now contain any character that is legal
'           for a file name, and can be any length.
'       8) Bubble point test now displays the bubble point pressure and
'           diameter before showing the end of test message box.
'       9) During initial bubble point, regulator won't increment more than
'           one click at a time until it has incremented at least 5 times.
'           This is to compensate for some machines where the first click
'           doesn't get very much flow but the second one does.  If you click
'           multiple times after the first click, you will overshoot the
'           target flow rate, and thus may overshoot the bubble point for
'           samples with large pores.
'       10) Added lines to turn off error trapping at several locations
'           which otherwise would have unexpectedly left error trapping on.
'       11) Fixed unit number display - on some windows it was being
'           overwritten by later calls to change the caption of the form.
'   6.54.35 4-8-99 to 5-6-99
'       1) Replaced common dialog file selector call to a more direct
'           call to the WinAPI function.  This both elminates the need
'           for the common dialog ocx in the installation set (and helps
'           to solve an incompatability with multiple versions of OCX
'           files over different versions of Visual Basic) and also
'           seems to solve a problem with the file selector box always
'           showing the full path in the file name box, which can be
'           confusing.  Calls to fsel now need to pass the handle of
'           the form that is doing the calling.  If a routine in a
'           module calls fsel, that routine needs to know the handle
'           of the form that called it.  Because of this, the routine
'           GetaFile now requires the handle of the form that calls it.
'           The simplest way to do this is to add me.hwnd to the call.
'           Using "Me" allows the current form to be renamed without having
'           to change any code.
'       2) While doing above, all form self-references are now changed
'           to using the "Me" word.
'       3) All code that handles units earlier than version 6 is now
'           commented out.  This code doesn't work with 32-bit mode anyway,
'           so it is just taking up space.  Since VB5 makes much larger
'           executable files anyway, it makes sense to preserve space.
'           Current code now expects that version=6.0 all the time.  It
'           is only checked on start up now.
'       4) Win32API calls now correctly updated so that they work with
'           Windows NT.  Some Integer parameters needed to be changed to
'           Long parameters to avoid overflow errors.
'       5) User list moved from capwin.ini to capusers.ini inside the default
'           user name.  The default user name must exist and can not be
'           deleted.  If the user list is not in the capusers.ini/defualt
'           then it is either moved from the capwin.ini file (to update
'           older users) or it is created.  This was done to avoid problems
'           with people moving the capwin.ini file (which contains all the
'           instrument hardware information) from one computer to another
'           and then finding that their user information is messed up.
'       6) Changed way user groups are deleted from the ini file - now using
'           one system call with vbNullString argument rather than defining
'           three different system calls depending on what type of argument
'           is used.  (vbNullString looks like a string to basic but looks
'           like a long integer of value 0 to the operating system.)
'       7) Optional "UnitNumber" entry in capwin.ini defaults to 0 and thus
'           doesn't show anything.  If other than 0, it will be added to the
'           title bar of every form to show which unit is being used.  This
'           makes it easier to run two units on the same computer and avoid
'           confusion.
'   6.54.34 3-26-99 to 4-1-99
'       1) Added readatcheck box and readat label to show status
'           of auxillary input bits (using the R@ command).  This
'           if only useful for custom systems that use the auxillary
'           input bits, and this check box is only visible if
'           the "READAT" entry in capstuff equals "Y".
'       2) Removed "extra low flow wet data" box - it now is always
'           turned off.  This was only needed in one specific
'           application for a user many years ago.  No current users
'           are using this feature, and if they turn it on by accident
'           it can mess up their normal results.  This box is now
'           replaced by a check box that turns on "advanced mode".
'           Advanced mode works like the previous version.  If you
'           are not in advanced mode, then a "lite" test setup screen
'           is presented first, only allowing you to change the output
'           file and sample ID.
'   6.54.33 3-11-99 to 3-23-99
'       1) Fixed problem with cv calibration where if you had to
'           retry the calibration the variable that keeps track
'           of the number of calibration points would be incorrect
'           leading to an "input past end" error message.
'       2) Added support for automated compression feature with
'           compression pressure gauge (high and low range),
'           motorized pressure regulator for compression pressure,
'           and compression actuator solenoid valve
'       3) Added compression pressure regulator calibration
'           Also - pressure regulator calibration (both types)
'           now will stop when the regulator reaches the maximum
'           count value or when the pressure gauge reaches full
'           scale, whichever comes first.  Both calibrations
'           use the same form, which determines which regulator
'           is to be calibrated by the compregcal variable, which
'           is true only if you are doing the compression regulator
'           calibration.  Also - if you abort the calibration it
'           will still save the table of how far you went.
'       4) Microflow volume can now be calculated on a machine
'           without a low flow meter - the first high flow meter
'           is used instead.
'
'   6.54.32 1-15-99
'       1) First vb5 32-bit version, being updated at the same time
'           as 6.54.30 16-bit version
'       2) Merged some support file functions, moved some files
'           into user directory, changed default directory to
'           c:\program files\capwin.
'           Capstuff.dat and board.loc files are now in the
'           capwin.ini file.  The user information that was in
'           the capwin.ini file is now in the capusers.ini file.
'       3) No long works with hardware version less than 6
'       4) All file access is through the FreeFile command
'       5) Now works with a 3-chamber bubble point tester.  If
'           CHAMBERS=3 then it will assume that you have the new
'           valves A, B, and C for chamber isolation.  Note that
'           if CHAMBERS=2 then this is for the 2-chamber perm-
'           porometer, which doesn't have isolation valves.
'   6.54.30 8-31-98
'       1) Fixed problem introduced in version 6.54.29 that would
'           set the first flow position of valve 2 incorrectly.
'       2) Apply +40 offset to valve 2 close limit to eliminate
'           backlash problem.  Also put -40 offset to open limit.
'       3) Added support for external frazier pressure gauge
'           if "FrazierPressureGauge" is "Y" in capstuff, the
'           microflow pressure gauge is replaced by an external
'           Frazier pressure gauge which is activated when the
'           maxpress parameter is set less than or equal to the
'           maximum pressure of the diffpg.  If there is a
'           Frazier pressure gauge, there is no microflow
'           venting valve.
'   6.54.29 6-5-98 to 8-31-98
'       1) Now works with new Bubble Point Tester - feature 32 with
'           high flow rate set to 0.  Doesn't have high flow meter.
'           If feature=288 then has solenoid valve 2, otherwise
'           if climit and olimit are 0, doesn't have any valve 2.
'       2) Program now compiled with "Option Explicit" to require
'           all variables to be declared.  This is being done to
'           prepare the way for moving to VB5-32.
'       3) Some variables changed from global to local or global
'           only to the form they are used in.
'       4) PulseDelay parameter from default parameters is now
'           used in elevated pressure liquid permeability to delay
'           from last increment so avoid false high flow readings
'           caused by the movement of the float during the rapid
'           increase in pressure.  (This didn't work properly in
'           internal version e and f, but was fixed in g)
'       5) Flow meter calibration now waits longer for stability
'           to give greater accuracy in bubble point.
'   6.54.28 3-17-98 to 5-8-98
'       1) Fixed pressure hold test so it can be aborted if the
'           pressure doesn't go up during initial pressurizing.
'       2) Fixed data averaging - all items on modify menu are now
'           called non-modally and they turn off their menu entries
'           while they are running.
'       3) Default version is now 6 so it doesn't mess up too much
'           when the capstuff file is corrupted.
'   6.54.27 1-5-98 to 3-9-98
'       1) Fixed problem in leak test when regtable is empty - this
'           would cause a "subscript out of range" error
'       2) If Olimit and Climit are the same, it won't crash.
'       3) At end of regulator calibration, if valve 2 or regulator
'           doesn't close all the way to zero, it will still end
'           correctly.
'   6.54.26 11-19-97 to 11-22-97
'       1) Changed some colors to make some screens easier to read
'       2) System now closes valve 1 when entering manual control
'       3) capstuff can now contain variable reg_pulse_min and
'           reg_pulse_max, only valid for motorized regulators.
'           These will default to 12 and 12 for same increment
'           throughout the range of the regulator.  Set to 12 and
'           4 (max=4) for older version operation.
'       4) Regulator calibration table now doesn't skip first
'           few values.
'   6.54.25 7-28-97 to 10-8-97
'       1) Added capability for having both penetrometer and microflow
'           differential pressure gauge in same machine.
'   6.54.24 7-11-97 to 7-25-97
'       1) Now allows fluid sensor reading in wet up/dry down test
'       2) Doesn't allow penetrometer to be pressurized above 200 PSIG
'       3) Motorized regulator now opens 100 counts less for initial bubble
'           point pressure (SBPP) to avoid overshooting.
'   6.54.23 5-22-97 to 6-19-97
'       1) At the beginning of the test it now waits for a stable reading on the flow meters
'       before it takes their zero point.  This helps if the test is started immediately after
'       running large flows through the flow meters and also helps in older 5.1 machines where
'       the regulator is always at least 1 PSI so the excercize_valve_2 routine can cause some
'       flow.
'       2) Bruce changed the wait from 10 (readtimes%), to 15 readings, to test for stability
'       for the flowmeters at the beginning of the test. 6-5-97 BMH.
'       3). Modification in PGCalibrate to properly handle the following case:
'           Only one absolute pressure gauge with range 250 PSI and two flow meter with range
'       30 cc/min and 500L/min. 6/5/97-97  by Jing Zhong
'   6.54.22 4-16-97 to 5-19-97
'       1) All demo modes should now work (they were actually fixed in a post-final
'       version of 6.54.21
'       2) "Temperature=2" added for dual-temperature liquid permeability testing.
'       3) On version 6 machines the feature number in the CAPSTUFF file is now checked
'       (if possible) and compared with the feature number of the instrument.  Only newer
'       instruments (4-1-97 and newer) have the capability of reading the feature number.
'       If the numbers differ, a warning is given but the machine is still usable.
'       4) Timeouts for autofill have been expanded and you can now purge an autofill system
'       5) if CAPSTUFF has CHAMBERS=2 then system will support having two sample chambers,
'       one for gas flow and one for liquid flow.  If this is the case, the end of the liquid
'       permeability test will change since you can't blow air through the bottom of the
'       liquid chamber in the same way as you can when one chamber is used for both.
'       6) If you abort while waiting for temperature to rise in elevated temperature liquid
'       permeability you will be given to option to continue with the test at the current
'       temperature or cancel the test.
'       7) Changed V2Percent flagging for versions 5.1 abd lower to allow exercising V2 very
'       small amounts. Useful for archeological style gas permeameters. The electronic regulator
'       does not zero. When v2 is exercised at beginning of test, it cannot allow flow, or the
'       flowmeter is set above the true zero unless V2PERCENT is around .5. BMH 5/7/97
'       8) V2INCR changed for versions 5.1 or less to allow small increments for gas
'       permeameters with small flowmeters. BMH 5/7/97
'       9) Changed the 75% flow limit warning messagebox for small flowmeters (such as in a
'       gas permeameter with only one 100cc flowmeter), 1000 cc or less, to allow full range
'       tests without the box popping up at the end.
'       10) Reformatted the Modify Parameter form to make it easier for customers to understand.
'       BMH 5/16/97
'       11) System now works properly with 1000 Torr absolute pressure gauge as the main gauge.
'       Before, it would give an error message that the instrument wasn't turned on.
'   6.54.21 3-27-97 to 4-14-97
'       1) Fixed pressure gauge calibration routine so it works with single gauge/
'       single high flow meter machines.
'       2) Remove integrity test option on machines that do not have integrity
'       flow meter.  (Older program would use low flow meter to simulate integrity
'       test and it didn't work very well.)
'       3) Fixed divide-by-zero error in microflow volume calibration if pressure
'       fails to increase (because of a leak or something like that).  It will now
'       display a message saying that the pressure is not rising and that there
'       is a possible leak.  If the leak is sealed and the pressure starts to rise
'       then the normal % and cc readings will be displayed.
'       4) Added support for "PMI Integrity Tester" which requires that the integrity
'       flow meter is installed (new variable integrity%) and that the high flow
'       meter is not installed (high and low range of high flow meter set to 0 cc/min).
'       This turns on new variable itester%, changes some of the captions, and
'       disables capflow and gas perm tests, leaving only integrity, bubble point
'       and pressure hold tests.  An Integrity Tester should not have an extra high
'       flow meter or penetrometer.  Adds the feature number addition of 256 which means
'       that valve 2 has been replaced by a solenoid valve and doesn't need to exercise
'       or anything like that.  This is new variable v2solenoid%.  Note that it is possible
'       to have an integrity tester with a motorized valve 2, but this would be a waste of
'       hardware.
'       5) Fixed problem with filename correction (added in 6.54.06) that would mess up if
'       your path contained a folder with an extension in its name.  Also updated file
'       selector so it could trap the user selecting a file name exactly the same as an existing
'       folder in the same path.
'       6) Parameters that are not used in a particular machine will not show up in the editor.
'       7) Square Pores is now not selected by default and you can turn it off if you turn it on.
'   6.54.20 1-26-97 to 3-17-97
'       1) Added support for temperature controller on auxillary input.
'       This is turned on by adding "TEMPERATURE=Y" to the capstuff
'       file.  Also add "TSX0=500", "TSX1=20500", "TSY0=0", and
'       "TSY1={whatever max really is}" and "TSUNIT={name of units}"
'       This will change the auxillary reading into a "Temperature"
'       reading in whatever units you specify.  It will also enable
'       the temperature output and allow the user to set the temperature
'       of the sample chamber for liquid permeability tests.
'       2) When autofill is refilling penetrometer it now gives a status
'       screen showing penetrometer value as it fills.
'       3) Fixed problem where when you selected a new type of test it would
'       then mess up the user selection box the next time you used it.
'       4) Fixed curve fit, editor, and averager to use secondary
'       fluid specification in liquid and gas permeability files.
'       They should also work with auxillary input fluid sensor files.
'       5) Added in capstuff variable "Compression=Y" for machines with
'       compression capabilities.
'   6.54.19 12-2-96 to 1-24-97
'       1) Fixed max liquid flow for older style autofills - pulsing of
'       drain valve after penetrometer fill was causing too much drop in
'       penetrometer on slow computers.
'       2) Changed captions on leak test - it shows absolute pressure but
'       put a "D" after the units, which not only was wrong, but was also
'       confusing for any unit other than PSI.  It now shows the pressure
'       unit and then " (Absolute)".
'       3) Added user-adjustable tortuosity factor
'       4) Disabled Execute and Calibrate menus when anything is currently
'       being executed or calibrated.  This stops two routines from trying
'       to control the instrument at the same time.  (They would step on each
'       other and cause problems.)
'   6.54.18 10-2-96 to 12-2-96
'       1) Added pressure gauge calibration routine by JZ
'       2) Added support for fluid sensor on auxillary input.
'       This is turned on by adding "FLUIDSENSOR=Y" to the capstuff
'       file.  Also add "FSX0=500", "FSX1=20500", "FSY0=0", and
'       "FSY1={whatever max really is}" and "FSUNIT={name of units}"
'       This will change the auxillary reading into a "Fluid Sensor"
'       reading in whatever units you specify.  It will also turn on
'       two new CFP tests for WUDU and DUWU with fluid sensor.  These
'       tests add some changes to the CFP file, so a new report program
'       will be required.
'       3) New Select_Test box with added options for microflow porometry
'       and nowait on wudu.
'       4) Alter Cover now uses rp_cover.txt in main directory rather than
'       one in user directory which wasn't working properly with current
'       report program.
'   6.54.17 8-1-96 to 9-12-96
'       1) Didn't add pressure gauge calibration routine by JZ
'       2) Removed upper limit on max. dist. between points in curve fit.  You
'           can now enter as large a number as you like.  (If your number is
'           greater than the maximum pressure in the test, it will have no
'           effect.)
'       3) New routine for "extra low flow wet data" test - resets regulator
'           after bubble point so that there is only small amount of flow through
'           low flow meter.  Old routine zeroed regulator, which would cause
'           long time delay while pressure built up again.
'       4) When you rename a group it now deletes the old group information.
'           It used to just leave the old group information in the INI file.
'       5) Caption changed in unit conversion box to give more information.
'       6) "End User" and "Test Reference" labels don't get changed to "line 1"
'           and "line 2" - they stay the way they should be.
'       7) New diff. pressure gauge based porometry tests using diff. pg in place
'           of a flow meter to measure very small flows.  Parameter file must
'           have maxflow rate of 1 cc/min to turn on this feature.
'       8) When running new diff. pressure gauge based test, status line now
'           shows differential pressure gauge reading and doesn't show flow meter
'           reading.  (The flow meter reading is ignored in this test.)
'       9) Added additional manual control for motorized pressure regulator
'       10) Now uses MSCOMM.VBX for communications with version 6 instruments
'           if you are using comm ports 1 through 4.  Will still work with PMI
'           comm board.
'       11) Turn, Abort, and Manual buttons only work when you are in hold.
'           Hold may be delayed in taking effect while some tasks that can't
'           be interrupted easily are going on.
'       12) Added support for features 64 (aux. analog input) and 128 (microflow
'           diffpg) in version 6.0.
'       13) Fixed problem on some version 6.0 machines where low pressure gauge
'           safety valve (valve 11) would be closed at the wrong time.
'       14) Added support for gas permeameter - same as capflow but with no
'           valve 1 or low flow meter, thus no bubble point or porometry tests.
'           You turn on this feature by setting the high range of the low flow
'           meter to 0.  (You also have to set the low range of the low flow meter
'           to -1 to shut off the integrity meter option.)
'       15) CV calibration now uses slower regulator increments if the maximum
'           flow rate is less than 50 l/min.
'       16) Will now work with CV table with only one value or a table where the
'           values keep increasing to the end.  Older program assumed that the
'           CV table would go up and then decrease near the end.
'   6.54.16 7-5-96
'       1) Added 0.1 second delay when pulsing valve closed to overcome race
'           condition in older version 5.3 small instrument board that would
'           allow valve 2 to close beyond the normal close limit when pulsed
'           on some instruments.
'       2) Fixed bug where if you edited the parameter file from the auto test
'           setup box and changed the name of the parameter file when you
'           stored it, the changed name would show up on the setup box but
'           the old name would be actually used for running the test.
'       3) Microflow test start prompt changed and protection added for microflow
'           pressure gauge in manual control.
'       4) Put MaxAirFlow routine back in cV calibration - it was left out of
'           version 6.54.15
'       5) MaxHighFlow variable holds original high range of regular high flow
'           meter, and this variable is used to compare to MaxFlow parameter to
'           determine which flow meter is used in the test.  This replaces the
'           comparison with the FY2 variable that holds the current high range
'           since this variable can change slightly during the test as the flow
'           meter is recalibrated and this can cause the wrong flow meter to be
'           used in the second half of a capflow porometry test in some cases.
'       6) Added preliminary demo mode for version 6
'       7) Version 6 hardware can now set the com port from the calib menu.
'       8) Added new "Wet up, Dry dn bp" test which will do a wet up, dry down
'           and automatically stop when the pressure down is less than the bubble
'           point (after a stable reading is taken).
'       9) "New User" copies linear and thickness units from current user
'       10) All capflow tests now have extended report info that says what type of
'           test was actually run.
'       11) Data points are stored with equilibrium coding added to the spacing
'           of the numbers.  The data points consist of a flow rate, a comma, and
'           a pressure value.  There is always a space before the flow rate.  There
'           is an optional space after the flow rate before the comma and an
'           optional space after the comma and before the pressure value.  The
'           coding is as follows:  A space before and after the comma means that
'           the first equilibrium routine was used.  No space before the comma
'           and a space after the comma means that the second equilibrium routine
'           was used.  A space before the comma and no space after the comma means
'           that the user forced the data point to be taken.  No spaces either
'           before or after the comma means that the alternative background
'           equilibrium routine was used.  Note that this is not used with some
'           of the special equilibrium options such as when aveiter=0.
'       12) Elevated pressure liquid permeability now asks for starting pressure.
'       13) Parameter file(s) used are now stored as extended values in data files
'       14) Deleting a group name now actually removes the group information from
'           the CAPWIN.INI file.  Before, it would remove the group name from the
'           list but all the information would still be taking up space in the
'           INI file.
'       15) Added message box to abort button to verify that you really mean it.
'   6.54.15 4-16-96
'       1) Added "Wet Up/Linear Dry" to capillary flow porometry type tests.  This will
'         run the wet curve normally and then extrapolate a linear dry curve based on
'         the maximum flow rates and pressures of the wet curve.  The dry curve will be
'         linear except for the constraint that it does not go below the wet curve.
'       2) CV pressure correction reformulated for greater accuracy at low pressures
'       3) New estimated bubble point pressure routine speeds test time
'       4) New Bloop helps maintain and achieve bublflow more precisely
'       5) New five_to_one variable in capstuff file now lets us know that the
'          Maximum attainable pressure is 30 PSIG.  This is in use for the leak test
'          and the autoparms setting of maxpres.  Also changed leak test such that
'          the user cannot enter a pressure over the range of the high pressure gauge
'          else and error message is returned and the leak test aborted
'          I think that the variable reg5% will be very usefull for future programmers
'          Who wish to optimize the code for low pressure samples etc.
'       6) Initial support for a user directory - used for rp_cover.txt report cover file
'       7) Support for version 6 feature number 32 (pneumatic regulator)
'   6.54.14 3-26-96
'       1) changed valve test routine from 5000 to 4500 cts.  This reflects the count
'         value that is the true maximum close limit and minimum open limit.  Some very
'         small travel valves can't get to 5000 but all valves must be able to get to
'         4500.
'       2) It now closes the low pressure gauge valve during BEFORE estimated bubble
'         point if the estimated pressure is greater than the range of the lowpg.
'       3) If the user saves a new parameter file in AutoParm, the AutoTest screen
'         will show that new file in either/both the Wet Curve and the Dry Curve.
'       4) Inserted a message box to ask the user which Curve parameter file to edit
'         if the 2 files are different and neither is checked.  (DWW)
'       5) Whenever the OK button is made visible on the Msgform, the focus is also
'         set to it.
'       6) the find volume routine has been rewritten
'       7) pressure generator increments for version 6 change the pulse width (cmd J)
'         depending on the current position.  At start they use 12.  At end they use 4.
'       8) graph in pressure hold test fixed
'       9) execercise_valve_2 timer modified - if the valve stops before the close
'         limit, it will retry four times before saying it is out of calibration.
'       10) Cleanout routine makes sure that extra pressure gauge is protected
'       11) Pressure regulator calibration (version 6) make sure extra pressure
'         pressure gauge is protected
'       12) Fixed rare case where low estimated bubble point pressure on machine
'         with only one pressure gauge would try to read nonexistant extra pressure
'         gauge.
'       13) New parameter in CAPSTUFF: PSIPERCC = PSI/MIN increase per CC/MIN flow
'         into system during hydrohead test.  If this parameter is present, the hydro-
'         head test will ask for rate of pressure increase.  If not present, it will
'         ask for the flow rate.
'       14) Series of changes made by JSD for version 6 machines to control better
'   6.54.13 2-5-96
'       1) New "HydroHead" test for Perm-Porometers - has new HydroHead% variable
'       2) Selection box cancel now works properly - returns T_Select$ as "Cancel"
'       3) I/O board calibration now works properly - it wasn't saving board.loc file
'         in proper place.  Variable ExePath was invalid - CAPWIN uses exe_path$
'       4) initial valve 2 testing should now be faster
'       5) Internal variables keep track of the last
'         state of the pressure regulator.  If you last zeroed the regulator,
'         it doesn't bother to zero it again at the beginning of the next test.
'         Note that this variable is cleared when you enter CAPWIN, so it will
'         always zero the regulator again the first time through.
'       6) Added second, stronger warning at end of liquid permeability test to
'         try to make sure that the user has drained the liquid from the sample
'         chamber.
'       7) Integrity test now works in version 6
'       8) New optional parameter in CAPSTUFF:  V2Percent = 0 to 100
'         This lists the percent open you want v2 for estimated bubble
'         point pressurization, integrity test, and elevated pressure
'         liquid permeability.  If you don't put anything, it is assumed
'         to be 100%.  Don't set this value too low or your tests may not
'         work!  This is only to speed things up a bit so you don't have
'         to wait as long for the valve to open all the way.
'       9) When you are asked to refill the penetrometer, you can now
'         click on "Stop Test" and abort the test.
'       10) Special debug mode added
'       11) Integrity test now asks for starting pressure
'       12) HydroHead test now asks for starting pressure
'       13) Bubltime changed to F/PT (average flow / pressure change * time)
'       14) minimum maxpressure parameter set from 1 to 0.
'       15) New optional parameter in CAPSTUFF:  3WayValve = Y
'         This means you have a 3-way protection valve on the low pressure gauge
'         and thus if the low pressure gauge is over pressure it doesn't need
'         to give an error until after it closes the 3-way valve to vent the
'         pressure.  Variable way3 set to true if this valve exists.
'       16) Timer doesn't start until actual testing starts
'   6.54.12 1-12-96
'       1) Fix display of penetrometer venting valve in manual control.  This was done incorrectly in
'          version 6.54.08.  The valve displays in the opposite manner from the actual valve position.
'       2) Added new "Air Bubble Purge" for non-autofill perm-porometers
'       3) Fixed initialization of Pass variable - caused problems on second run
'       4) Fixed parameter editor dirty flag - won't tell you things have changed if they haven't.
'       5) New Parallel stability mode - if readings are unstable but aren't going anywhere, forces
'          reading after 30 seconds.
'       6) Regulator calibration for version 6
'   6.54.11     1)  Removed Add and Remove Test
'               2)  Corrected a couple of spelling mistakes
'               3)  In AutoParm Form, added the string to be saved to the MsgBox statement for saving
'               4)  Corrected glitch in AutoTest screen whereby changing the user would not update all
'                   features of the AutoTest screen.
'               5)  Added feature which allows the user to change a group name (Its useful for when the
'                   technicians spell a group name incorrectly).
'               DWW 1/10/96
'   6.54.10     1)  On the Progress Form, displayed both the filename and the sample ID above the graph.
'               2)  Removed Microflow Analysis from Add/Remove Test from fear of user confusion.  The group felt that
'                   user would add the test, not knowing that you needed an entirely different machine.
'               3)  Added a feature which allows the user to change the user from within the Autotest form.
'                   When this is done, it also updates the necessary components on the main screen.
'               DWW 1/4/96
'   6.54.09     Added standard Small Instrument Board Calibration Module (Devin Sundaram, 12/29/95)
'   6.54.08    In version 6, the regulator now moves based on the pulsewidth and the capcal.d8a file for version 6
'               stores regulator counts rather than clicks.  Since no version 6 machines have been shipped yet, this
'               change in protocol will not affect customers.
'              Also added new Manual Control with shapes in place of a painted drawing - this makes updates easier.
'               New manual control also is more multi-tasking "friendly".
'              Also fixed slight problem in diffperm volume calibration.
'              Protection for low pressure gauge is now fixed - valve will automatically shut off the gauge when you are in manual
'               control if the pressure goes over 23000 counts, and valve will not open if the pressure on the main pressure gauge
'               is over the top pressure rating of the low pressure gauge.
'   6.54.07     1)  Changed appearance of Title Screen.  2)  Added feature in Modify Menu which allows user to add any type of
'               non-standard CapWin Test.  Thus far, there are only 2 such tests, Microflow and Square Pore Analysis.  3)  Added PC
'               board calibration to calibrate menu.  (DWW)
'              Also changed how high low pressure gauge could go without causing error during test.
'   6.54.06     Corrected file name entry problem.  In TestScrn, there were no checks on a correct filename, hence
'               ending a filename with "." or ".bla" was acceptable.  Filenames now end with ".cft" regardless of user entry. (DWW)
'   6.54.05    Added test of low pressure gauge during the test to make sure that the valve is not leaking.
'              Also added way of reading low pressure gauge while leaving the valve closed.
'              Manual control now shows correct position of low pressure valve at all times.
'              In manual control added shift-I and shift-D to increment and decrement by 10
'               Added new machine type with 'DIFFPG="Y"' in Capstuff file for adding a low pressure gauge after the
'               sample chamber.  Valve 14 now drains this gauge.
'               New type of test - diffusion permeability
'              Added linear_unit_name$, linear_unit_conversion# (to cm), thick_unit_name$, and thick_unit_conversion#
'   6.54.04     Modified regular equilibrium routine so that during wet up it will not let the
'               pressure fall back.  (Similar problem to that solved in 6.53.30, but seen less often)
'               Also changed penetrometer autofill routine to compensate for transition point between offscale and
'               in-scale values.  Some penetrometers have a small region where incorrect seeminly in-scale values
'               are read.
'               Also added capstuff parameter "TopFill" which is equal to "Y" when you have a new top filling autofill
'   6.54.03     Changed beginning of ReadXReturnX4 to see if it would stop valve 1 from making noise when valve 2 pulses
'               during the auto test.  (V1 doesn't make noise if valve 2 is pulsed during manual control)
'   6.54.02     Modified routine during wet curve when pressure falls.
'              Also fixed potential problem when using 5000 Torr gauge with 1000 Torr gauge because high range of 1000
'               Torr gauge is same as (or higher if 1000 Torr is differential) than 5000 Torr gauge low range which could
'               mess up gauge recalibration routine.
'   6.54.01     Rewrote gauge calibration during auto test (again!)
'               This version remembers counts at atmospheric pressure, and atmospheric pressure, and whenever
'               a gauge's conversion values are changed it makes sure that it would still read atmospheric pressure
'               if the count value was set back to what it was at atmospheric pressure.
'               If this works, it will be used later in a seperate gauge calibration routine.
'              Also, "Record" button is ignored within one second of reading a value.  This will stop the program
'               from taking a data point too soon.  What can happen is that the user presses the Record button,
'               thinking that the point was taking too long to reach stability when in fact it had just reached
'               stability and was storing the result to the disk file.  The "Record" button is then used to force
'               the NEXT data point will then be taken immediately, not allowing any time for the reading to stabilize.
'   6.54.00     First version that works with new version 6 hardware (George based)
'   6.53.31     Removed "Something is wrong with data" message if you are running just a bubble point.
'   6.53.30     Modified new equilibrium routine (AveIter=0) so that during wet up it will not let the pressure fall
'               back.  (On some samples, the pressure rises, pores are opened, and then the increased flow causes the
'               pressure to drop, leading to improper results.)
'               Also added menu item to turn off the advanced low flow readings.
'               Also fixed initial scale of test screen.
'               Also changed pressure gauge re-calibration so that when you switch from the low to high pressure
'               gauges, it does an offset shift, and when you switch from low to high range on the same gauge, it
'               does a scale shift.
'   6.53.29     Changed gauge re-calibration when changing range during an automated test.  It now shifts the Y2 value
'               as opposed to the Y1 value, and if the Y2 value is being told to shift by more than 10%, it doesn't do
'               it (it assumes that it is in error)
'               Fixed file selector call in data editor for saving permeability curve - I/O flag was set wrong so you
'               couldn't save to a new file that didn't already exist.
'   6.53.28     Modified section of gauge reading routine where range switches.  It used to read low range and then
'               read the high range and reset the high reading to equal the low reading.  It now reads the low, then
'               high, then low again and makes the high reading equal to the average of the two low readings.
'               This should make the crossover smoother when the flow is changing rapidly.
'               Also changed bubltime value for offscale readings (when pressure didn't go up) from 1000 to 99999.
'               Also changed new stability routine (when AveIter=0) so it looks at both phase shifts in pressure as
'               well as flow (old new routine only looked at flow phase shifts)
'   6.53.27     Rewrite of gas selector for gas or liquid permeability test.  If not one of the preset values,
'               has lines for gas name and gas viscosity.
'   6.53.26     Made sure SHFP was at least 3 clicks on the regulator
'               Dry down curve can now end when valve 2 is within 1% of close limit.
'               Also fixed file selector for when original path is no longer valid, such as when you
'               store a file to the A disk and then don't have the A disk in the drive when you recall
'               the file selector (it will switch to C), or if the path you last used is no longer valid
'               (it will go up one level until it finds a valid path or reaches the root, and if the root
'               is not valid it will find another root that is valid.)
'               Also reads reads high flow meter once before moving valve 2 so correct high flow meter is
'               selected earlier to allow it to settle in better.
'               Also fixed some lockups at end of dry down curve where valve 2 would not close all the way
'               to the close limit because it was being pulsed very slowly.  Now if the valve doesn't respond
'               in 10 seconds, it will stop trying to move it.  Also, if valve gets to within 1% of close limit
'               (1% of distance from close to open limit) the valve will be called closed.
'   6.53.25     New expanded bubble point routine.
'               Also increased shown resolution of PulseWidth parameter.
'               Added new routine to keep NOWHERE% counter at 2 if going up and either P or F was actually improving.
'               Changed SHFP definition from 2 cc/min low flow to .2 cc/min low flow.
'               Also, when bubble point is done, it zeroes regulator and then sets reg to SHFP.
'               This should allow lower flow rates for initial high flow section of test.
'   6.53.24     Fixed display of maximum pressure in pressure hold test - it now shows correct pressure for non-PSI units.
'   6.53.23     New file selector
'               Improved aborting from leaktest
'               During special fast test, pressure regulator is incremented only once until some real data point
'               is actually taken.
'               New pressure regulator increment factor that increases number of clicks of pressure regulator as
'               test goes on to try to achieve even distribution of data points.
'               Maximum PREGINC raised from 10 to 30.
'               Rewrote end of liquid permeability test to try to avoid getting fluid in the flow meters
'               Definition of Integrity meter changed.  Now if low range of low flow is more than 0.5 of high range
'               then it is an integrity meter.  Before, low range had to be more than high range.
'               New initial fill routine for autofill of penetrometer.  When stops filling, pulses drain valve to
'               relieve pressure below sample.  Also, stops early to see how much it goes over stopping point and
'               then resumes and uses this information to determine when to stop to avoid overfilling.
'               When they create a new group, that group is selected, but when they exit and re-enter, it doesn't
'               remember that this group was selected.  This is now fixed.
'               If they try to run the machine when the main pressure gauge is reading under 5000 counts on the low range
'               the program will warn them that their instrument may not be turned on.
'               Also fixed a few cases where mouse pointer stayed in busy mode.  May not have all of them yet.
'   6.53.22     Fixed auto-test setup screen initialization.  (When last test was liq. perm, it shows
'                   the line about surface tension when it is not necessary.)
'               Fixed: if you abort during first pass bubble point, second pass doesn't start up and then abort
'               Fixed: You can abort from bubble point if your estimate is too high
'               Fixed: Printout from auto parms editor matches screen layout (added new parameter)
'               New: CV calculation records how many regulator clicks were used to determine CV, and if
'                   the user tells the program to retry, it starts at that many clicks to start with.
'               Also: On retry of CV, program used to display incorrect old CV value.  This is fixed.
'   6.53.21     New bubble point options for large filters
'               also made turnaround flag seperate from other flags so that you can click on the
'                   turnaround flag and then on the record flag and force it to take the next point
'                   and then turn around.  (Before, they both used the same variable, so clicking
'                   on record would overwrite the turnaround command.)
'               added minbppres parameter.  Won't take bubble point at a pressure below minbppres
'               also changed maxliqflow test so that if the penetrometer doesn't fill at all during
'                   the first 30 seconds, when you tell it to retry, it will go back to the beginning
'                   of the fill procedure, not to the middle as if the penetrometer is somewhat filled.
'   6.53.20     Traps lockup that happened when capcal.d8a file was corrupted or missing.
'   6.53.19     fixed wet up dry down test so that if you end the wet up because of maximum flow
'                   or pressure, when you start coming down you don't trigger if the flow or
'                   pressure is still over the maximum.
'               also fixed liquid permeability display so screen doesn't scroll back to top with
'                   every data point.
'               also changed manual liquid fill display to show you the counts all during fill,
'                   even after the program thinks you are done.
'               also added constantly updated status line to liquid permeability to show liquid
'                   height, pressure, and time
'               also added lots of DoEvents inside of timing loops
'   6.53.18     New parameter CFANAL which can equal "Y" in capstuff file to change name of
'                   captions on main screen
'   6.53.17     Updated high-flow bubble point so it works the same as the low-flow bubble
'                   point routine.
'   6.53.16     new routine when aveiter=0
'                   read flow twice and note direction it is going in
'                   Wait until it changes direction eqiter number of times,
'                   or until time elapsed=mineqtime
'                   By setting eqiter=0 or mineqtime=0 you wait no time at all
'                   by setting eqiter=999999 you wait mineqtime for each point
'                   by setting mineqtime=99999 you always wait for eqiter direction changes
'               Also allows zero flow to be stored to data files.  Before, it would
'                   store 1 in place of zero or negative values, and this messed up things
'                   with very low flow samples.  CAPREP has also been modified to allow
'                   gas permeability files to have zero flows in them.
'               Also fixed typeing error which caused eqiter parameter to remain as it was
'                   in the default parameter file and not be updated to the value it was set
'                   in the users desired parameter file for an automated test.  All other
'                   user desired parameters were updated correctly, but eqiter remained as it
'                   was set in the default parameter file.  This typeing error originated
'                   in version 6.0 or 6.1 as it was not in the GWBASIC version but was in the
'                   last QBasic version and seems to have migrated into the windows version.
'               Also, eqiter and aveiter are now allowed to go to 0 in parameter editor
'               Also fixed starting flow on screen graph to correctly reflect startf in users
'                   testing parameter file.
'               Also changed system so that valve select lines and direction line all go back
'                   to zero when no valve is actually moving.  This happens whenever a flow
'                   meter or pressure gauge is read and no valve is currently energized.
'                   This is meant to fix a strange bug that causes the high flow meter to
'                   shift its zero point depending on the status of the direction line.  The
'                   valve status lines are zeroed just in case they cause problems.  The
'                   actual hardware cause of this problem has not been determined.
'   6.53.15     fix time1k calling problem in manual control for pulse of valve 2
'   6.53.14     adds new "Fast Regulator" mode with valve percentage setting
'   6.53.13     turns on demo mode if board.loc = 0 (PA=0)
'   6.53.12     starts using newbet.dll again due to problem with truetype fonts
'               doesn't need time1k as that is now done internally
'   6.53.11     started using pmi.dll which is compiled with GFA Basic for Windows
'               no longer needs BET.DLL or NEWBET.DLL
'               PMI.DLL also has 1/1000 sec timer (TIME1K) so no need for procsped.d8a
'   6.53.10     added help links (corrected) and display penetrometer readings
'               during refill.  Fixed penetrometer fill waiting so that liquid
'               must go up to half-way point before it will terminate when count
'               goes off scale (previously, if you filled slightly and then let
'               the level drop so it was off scale on the bottom end it thought
'               that the penetrometer was filled).
'               Also fixed display of pressures during leak test for non-PSI units
'               and user input of maximum pressure during liq. perm. to differential.
'   6.53.09     fixed bug in p0first correction
'   6.53.08     fixed bug in "About Capwin"
'   6.53.07     added str$ to all print # commands
'   6.53.06     Removed parameter "Special1".  It's function can be duplicated
'               by setting AveIter to 0.  Added parameter "PulseDelay" in it's
'               place, which is a delay time after each increment of the pressure
'               regulator during bubble point.
'   6.53.05     Added warnings for off scale low pressure gauge and tighter
'               criteria before doing extra interpolated data points
'   6.53.04     Fixed problems when using version 5.2 instruments
'   6.53.03     Moved "Refill Penetrometer" message after zeroing pressure
'   6.53.02     cleaned up Reply% initialization for some msgboxes
'   6.53.01     cleaned up some things in test selection box
'   6.53.00     converted to visual basic version 3.00
'   6.52.06     1.  Fixed hollow sample stuff
'               Code taken over by Ron V. Webber (ym)
'   6.52.05     1.  Support for hollow samples added    change by ym
'                   to data file, test setup, and all
'                   other programs.  In data format,
'                   if diameter read in is 0, new next
'                   line contains Diam, Cyl_Len
'                   otherwise Cyl_Len is set to 0.
'
'   6.52.04     1.  Msg that V2 needs to be calibrated  change by JP
'                   now aborts test.
'
'   6.52.03     1.  changed penetrometer messages for   change by JP
'                   better readability
'               2.  text inputs for auto-test setup     change by JP
'                   now are highlit for easier editing
'   6.52.02     1.  bug fixed to allow leaktest to run
'                   beyond midnight                     change by SB
'               2.  timer-bar added to leaktest wait    change by JP
'   6.52.01     1.  "Credits" text added                change by JP
'               2.  leaktest changed to diff. pressure  change by JP
'   6.51        updated Capwin w/ data averager...      changes by MM
'   6.5         original windows CapWin                 changes by MM
'   6.1         qb45 seperate lists wet dry             changes by JP
'
'   6.0         first qb45                              converted by JP
'   < 6.0       written in Basica                       created by Ron V. Webber (ym@lightlink.com)

'   Hardware Ver.
'   < 4.0 Atari controller / interface boards.  Will not work with these versions
'   4.1 through 4.3 same as 5.1 through 5.3 using version 1 IIB with 100-4000 ADC
'   This program requires version 5.1 or higher hardware.
'   5.1 norgren regulator with version 2 IIB with 500-20500 ADC
'   5.2 internal controller 8 valves
'   5.3 internal controller 16 valves
'   5.31 16 valves and 2 high flow meters
'   6.0 George based, with new "Feature Number"
'   7.0 George based, with 16-bit ADC and I/P converter standard
'
' Special_Factor was removed in version 6.54.21g - it is redundant with t_factor
' END DOCUMENTATION

