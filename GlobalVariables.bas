Attribute VB_Name = "GlobalVariables"
Option Explicit
''''''''''''''''''
'8/30/2012
'Sam Bouabane and Sean Vesley
'These values are used by the pressure hold test (CAPFLOW.BAS)
'They are set in prefsForm when the corresponding box is filled out then loses focus
Global startThresholdMin As Single
Global startThresholdMax As Single
''''''''''''''''''
Global AirResistivity As Boolean
'
Global testInt As Integer

'
''''''''''''''''''
'V2 Limits fix
Global skipLimits As Boolean
''''''''''''''''''
'latching valves
Global LatchValves As Boolean
Global LatchingValves(23) As Boolean
''''''''''''''''''
Global tempProbe As Boolean
Global tempProbeMinC As Long
Global tempProbeMaxC As Long
Global tempProbeMinV As Long
Global tempProbeMaxV As Long
Global tempProbeChannel As Integer
Global tempProbeLabel As String
''''''''''''''''''''''''''''''''''
Global BPPostPurge As Boolean
Global BPPostPurgeCounts As Integer
Global BPPostPurgeDuration As Integer
'''''''''''''''''''''''''''''''''
Global geoFlowMeterSwitch As Boolean
Global geoValveClosed As Boolean
Global geoIncrease As Integer
Global geoRegCounts As Integer
Global GeoExtraValve() As Integer
Global geoPoreValve As Boolean
Global geoFlow As Long ' add by jglann for pausing test when flow reaches some value cc/L
Global openValve As Boolean
Global closeValve As Boolean
Global stopValve As Boolean
Global pulseValveClosed As Boolean
Global pulseValveOpen As Boolean
Global closingValve As Boolean
Global closedPercent As Double
'''''''''''''''''''''''''''''''''''
Global pnumValve As Integer
Global PneumaticMotor As Boolean


Global SkipFlowStabilize As Boolean
Global S_Version As String    'current software version
Global Const ProgAbbv = "Cap"
Private dw As New frmDebugWindow

'Global Const Special_Factor = 1    ' use 0.7 for Norton, 1 for everyone else

Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Const SW_SHOW = 1
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long
Declare Function WPPS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function GPPS Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'The following information is for User Account Control
Public Type UAC_User
    Username As String
    Password As String
    accessLevel As Integer
End Type

Global UAC_Enabled As Boolean
Global UAC_UserCount As Integer
Global UAC_Users() As UAC_User
Global UAC_LoggedIn As Boolean
Global UAC_CurrentUser As Integer
'Global UAC_Crypt As vbCrypt.EncryptionTools

' capwin_feature_number now stores this feature number for comparison
' purposes only.  Do not use this to determine what features are present
' outside of the capwin.ini reading routine as this will make it harder to
' add new features or shift the feature bits in future hardware versions
Global capwin_feature_number As Integer

Global Servo_Table_Exists As Boolean

'Added 4-23-10 rvw version 6.72.017 for Rice Univ. with Resin intrusion
 ' Resin_Diverter_Valve is set to 0 if there is no Resin system
 ' If there is a Resin system, Resin_Diverter_Valve is set to the valve number of the valve
 Global Resin_Diverter_Valve As Integer
 Global Resin_Fill_Height As Double
 Global Resin_Start_Height As Double
 Global Resin_Drain_Seconds As Double
 Global Resin_Start_Pressure As Single
 Global Resin_Increment_Pressure As Single
 Global Resin_Number_Points As Integer
 Global Resin_Stable_Seconds As Single
 
'Added 4-9-10 rvw
Global Num_Microflow_Volumes As Integer
Global Microflow_Volume2_Valve As Integer
Global Microflow_Volume3_Valve As Integer
Global Current_Microflow_Volume_Index As Integer
Global microFlowUseAllVolumes As Boolean

'Moved here 8-23-10 JDF
Global lperm_startp As Double
Global lperm_maxp As Double
Global lperm_stepp As Double
Global lperm_maxwait As Single
Global lperm_maxpoints As Integer

Global lperm_v6_reginccount As Integer
Global lperm_v6_regincwait As Integer

Global lperm_init_wait As Integer
Global lperm_user_cancelled As Boolean

'Added 10-30-09 - AJB
Global num_sample_pressure_gauges As Single
Global safe_temperature As Single
Global cartridge_tester As Boolean
Global cartridge_tester_side As Integer
Global currentPressureGauge As Integer
Global switchLPGauge As Boolean
Global testing As Boolean
Global testGasStatus As Boolean
Global depressurizeBeforeTest As Boolean
Global savePreBPdata As Boolean
Global topDownLp As Boolean
Global currentlySelectedChamber As Integer
Global tank_level_exists As Boolean
Global min_tank_fill_level As Single
Global tank_level_location As Integer
Global minFPT As Single
Global test_piston As Boolean
Global hasMultipleMVs As Boolean 'Used to be hasSecondMV2
Global useBubblerMV As Boolean 'flag to set control in manual control to the humidity bubbler motor valve
Global bubblerMV_OLIMIT As Long 'Open / Close limits for the bubbler motor valve
Global bubblerMV_CLIMIT As Long
Global motorValveIndex As Integer
Global numberOfMotorValves As Integer
Global motorValveSwitchFlow As Long
Global number_of_wetting_valves As Integer
Global wetting_valve(3) As Integer
Global current_wetting_valve As Integer
Global wetting_valves_latch As Boolean
Global no_chamber_bypass As Boolean
'Global mv2_reg_pos As Long
Global mv2_start_pos As Long
'Global mv3_reg_pos As Long
Global mv3_start_pos As Long
Global mv1RegEndPos As Long
Global mv1EndPos As Long
Global mv2RegEndPos As Long
Global mv2EndPos As Long
Global monitorFlow As Boolean

Global mv1_index_char As String
Global mv2_index_char As String
Global mv3_index_char As String

' 11-10-08
Global number_of_pistons As Integer
Global piston_valve(3) As Integer

Global ztempPressure$, ztempFlow$, ztempDiam$

'Added 11-21-07 --Denis
Global auto_wet_enable As Boolean ' true if the hardware exists
Global auto_soak_enable As Boolean ' true if older wetting hardware exists (without pump)
Global rotating_chamber_enable As Boolean ' true if we have motor(s) on chamber(s) to rotate sample
Global auto_wet_used As Boolean ' true if they want to use it
Global auto_wet_wet_time As Single ' time, in seconds, for wetting
Global auto_wet_volume As Single 'volume to wet the sample
Global use_auto_wet_volume As Boolean
Global auto_wet_soak_time As Single ' time, in seconds, for soaking
Global auto_wet_drain_time As Single ' time, in seconds, for draining
Global auto_wet_reverse_time As Single ' time, in seconds, for reversing wetting pump
Global auto_wet_pump_speed As Single ' value 0-255 determines how fast the pump forces liquid through
Global auto_wet_fill_height As Single
Global auto_wet_rotating_speed As Single
'added 11-29-07 --Denis
Global chamber_ready As Integer
Global chamber_selected1 As Integer
Global chamber_selected2 As Integer
Global sequentialTesting As Boolean
Global tempChamberVal As Integer
Global tempChamberValB As Integer
Global fileIncreaseNum As Integer

'added 12/11/07 --Denis
Global Fill_ValveA As Integer        'The fill valve number from Capstuff
Global Drain_ValveA As Integer       'The drain valve number from Capstuff
Global Fill_ValveB As Integer        'The fill valve number from Capstuff 2nd Chamber
Global Drain_ValveB As Integer       'The drain valve number from Capstuff 2nd Chamber
'Global SecondPiston As Integer       'The valve number of the 2ch piston for 2nd Chamber

Global PerformanceFrequencyValue As LARGE_INTEGER
Global elev_lqperm_ck As Boolean
Global elev_lqperm_startPres As Boolean
Global elev_lqperm_steppres As Boolean
Global elev_lqperm_maxpres As Boolean
Global elev_LqPerm_wait_sec As Boolean
Global elev_lqperm_maxpoints As Boolean
' MsgBox return values
Global Const Help_Index = 3
Global Const Help_Quit = 2
Global Const Help_Context = 1
Global Const MAXVALS% = 22 ' Number of parameter values in autoparm

'The following variables tell what features are present
'in the instrument - this is stored in the capstuff file
Global version As Currency, H2OPERM As Boolean, ExtraPG As Boolean
Global ver1or2 As Byte, ver2or3 As Byte, ver1or3 As Byte, ver1or20 As Byte
Global CFAnal As Boolean
Global DiffPG As Boolean
Global v20_exists As Boolean
Global WESA_enabled As Boolean, WESA_exclusive As Boolean
Global serial_number$, PDrop As Boolean
Global hydrohead As Boolean
'Global hhAsBurst As Boolean ' this is getting confusing
Global burst As Boolean
Global mullen As Boolean
' 6.71.38x new variable for hydrohead tester
Global hydrohead_exclusive As Boolean
Global TopFill As Boolean, compression As Boolean
Global allowZeroCompression As Boolean
Global integrity As Boolean, itester As Boolean
' 6.71.67 temperature% is replaced by new variables (see below)
'Global temperature%
Global fluidsensor As Boolean
Global GasPerm As Boolean, v2solenoid As Boolean
Global FrazierTester As Boolean
Global recirculation As Boolean
Global doorlock As Boolean
Global flow_status As Integer
Global piston_status As Integer
Global manual_aux_click As Integer ' tells which of multiple auxillary things were last clicked
' feature% has been removed from the global list - it is now local
' to the procedure that reads in the capwin.ini file
Global chambers As Integer, newreg As Boolean, BPTester As Boolean
Global m_bBPCreateLogFile As Boolean
Global m_bBPFindingForCartridge As Boolean
Global m_nBPPressureArraySize As Integer
Global m_bDeviceIsReadyForDrainAndFill As Boolean
Global dry_chambers As Integer
Global multiChamberSystem As Boolean ' only true if there are multiple chamber isolation valves
Global manualMultiChamber As Boolean ' only true if we have to manually switch chambers
Global manuallySelectedChamber As Integer
Global liqpermonly As Boolean
Global microflowporometer As Boolean
Global Auto_fill As Boolean, SqrPore As Boolean, xhflow As Boolean
Global xhflow_meters As Integer
Global nov2 As Boolean
Global NoFailSafe As Boolean
Global te_number%, FrazierPressureGauge As Boolean
Global reg5 As Boolean, way3 As Boolean
Global auxin As Boolean, dpgplus%
Global autocompress As Boolean
Global safetyup As Boolean, safetydown As Boolean
Global safetyupdoor As Boolean, safetydowndoor As Boolean
Global safety_canceled As Boolean
Global autopiston As Boolean ' can't have both this and autocompress
Global readatenabled As Boolean
Global dualregulator As Boolean
Global externalhydrohead As Boolean
Global Second_Penetrometer As Boolean ' if they have a second penetrometer
Global penetrometer_select As Integer ' which penetrometer they have selected (1 or 2)
Global Second_Penetrometer_V9 As Integer ' 9 or 19 (for old system)
Global Second_Penetrometer_V12 As Integer ' isolation valve for second system
Global Second_Penetrometer_V13 As Integer ' fill valve for second system
Global Second_Penetrometer_V23 As Integer ' extra drain valve
Global AirTop As Boolean ' if air comes to the chamber from the top
Global Drain12 As Boolean ' if there is a drain valve (12) in the system
Global Drain12Motorized As Boolean ' if drain valve 12 is motorized
Global hhRunAsMullen() As Boolean  ' determines if the hydrohead test is a Mullen test
Global hhRunAsBurst() As Boolean  ' determines if the hydrohead test is a Burst test
Global hhRunAsHydrohead() As Boolean  ' determines if the hydrohead test is a real HydroHead test

Global sampleEjectionSystem As String

'The following variables are for hardcoding which tests run on which penetrometer sides
Global liqperm_penetrometer As Integer
Global hydrohead_penetrometer As Integer
Global mullen_penetrometer As Integer
Global burst_penetrometer As Integer

Global reg1pmax As Single
'Global special_ambient As Boolean ' if they are running an "ambient" test on a lowered penetrometer
Global supervisor As Boolean
Global superpass$
Global lvperm_enable As Boolean
Global lvperm_exclusive As Boolean
Global lvperm_numvalves As Integer
Global debug_button_enable As Boolean
Global ip_reg_enable As Boolean
Global ip_creg_enable As Boolean
Global low_flow_controller As Boolean
'Global crossoverdebug As Boolean
Global system_font As String
Global font_size As Single
Global font_bold As Boolean
Global vacuum_purge_enable As Boolean
Global num_vacuum_purge_cycles As Integer
Global just_did_vacuum_purge As Boolean
Global bubbler_enable As Boolean ' true if there is a bubbler attachment
Global bubbler_selected As Boolean ' true if they want to use the bubbler
Global BubblerLevelChannel As Integer
Global BubblerLevelZero As Long
Global bubblerLevelSpan As Long
Global v22_exists As Boolean ' true for ballard, with the extra recirculation valve 22
Global reg2_high_flow_switch_count As Integer
Global status_lights_enable As Boolean
Global status_lights_value As Integer ' 0 for off, 1 for red, 2 for yellow
Global piston_position_transducer_exists As Boolean '6.71.123.01
Global slurry_tube_exists As Boolean                '6.71.123.01

'The following variables hold calibration information for
'the various transducers - from capstuff file
Global FX1(2, 5) As Long                                ' Zero Count Value
Global FX2(2, 5) As Long                                ' Full Count Value
Global FY1(2, 5) As Single                              ' Real Zero Value
Global FY2(2, 5) As Single                              ' Real Full Value

Global disableSpanAdjustments As Boolean

'6.71.123.04 Global PX1(7) As Long, PX2(7) As Long
'6.71.123.04 Global PY1(7) As Single, PY2(7) As Single
Global PX1(11) As Long, PX2(11) As Long '6.71.123.04
Global PY1(11) As Single, PY2(11) As Single '6.71.123.04
Global fsx0 As Long, fsx1 As Long
Global fsy0 As Single, fsy1 As Single
Global tsx0 As Long, tsx1 As Long
Global tsy0 As Single, tsy1 As Single
' pen500 and pen20500 are left named as they are in the capwin.ini
' file even though they may not actually be representative of 500 and
' 20500 counts any more.
Global PEN500 As Single, PEN20500 As Single                 '(Tim Richards note mark)
Global PSIPERCM As Single, CSECAREA As Single               '(Tim Richards note mark)
Global P2PEN500 As Single, P2PEN20500 As Single
Global P2PSIPERCM As Single, P2CSECAREA As Single
Global PENZERO As Long, PENTWO As Long, PENSPAN As Long
Global P2PENZERO As Long, P2PENTWO As Long, P2PENSPAN As Long
Global oLimit As Long, cLimit As Long, reg_cl As Long, reg_ol As Long
Global olimit2 As Long, CLIMIT2 As Long
Global olimit3 As Long, CLIMIT3 As Long
Global creg_cl As Long, creg_ol As Long
Global MaxAirFlow As Single, MaxLQFlow(2) As Single, liquid_lohm(2) As Single
Global PSIPERCC As Single, Diff_Volume(10) As Single
Global cv As Double ' actually now taken from a table
Global V2Percent As Single
Global reg_pulse_min As Integer
Global reg_pulse_max As Integer
Global MaxHighFlow As Single
Global aux_p1_span As Single
Global aux_p2_span As Single
Global max_liq_pres As Single
Global lv_valve_pulse_timing As Single
Global piston_area As Double ' cross sectional area of compression piston
Global fixed_sample_diameter_cm As Double '6.71.123.14
Global use_fixed_sample_diameter_cm As Boolean '6.71.123.14

'The following variables keep the unit name and conversion
Global fsunit$, tsunit$
Global PCNV As Double, PU$
Global linear_unit_name$, linear_unit_conversion#
Global thick_unit_name$, thick_unit_conversion#

Global cel1(100) As Single
Global cel2(100) As Integer
Global cel3(100) As Byte
Global cel_i As Integer
'The following variables keep track of hardware locations
' based on the board.loc file
Global PA As Integer
' ComLoc% is used to determine demo mode as well
Global ComLoc%
' these are not used any more - they were used in version 5
' ANBIT%, ORBIT%, ComCheck%
'Global open_flag%, valve_disable%, valve_enable%

'The following variables hold the current parameter file
Global AVEITER As Single
Global BUBLFLOW As Single
Global BUBLTIME As Single
Global EQITER As Single
Global flowslew As Single
Global PulseDelay As Single
Global Maxpres As Single
Global MAXFLOW As Single
Global mineqtime As Single
Global MAXPDIF As Single
Global MAXFDIF As Single
Global PULSEWIDTH As Single
Global PRESSLEW As Single ' (can be non-integer since in version 7 it represents counts * 3)
Global preginc As Single '(can be non-integer in some cases)
Global STARTP As Single
Global STARTF As Single
Global V2INCR As Single
Global ZEROTIME As Single
Global minbppres As Single

'The following variables hold status information about hardware
' such as valve status, transducer range, etc.
Global FUSE%, DPress%, vflow%, CPress%, EPress%
'6.71.123.17 Global Vpos(-8 To 25) As Byte, HFLOW%, lflow%, Pres%.  02/09/16 added tempPres%, firstRun
Global Vpos(-8 To 45) As Byte, HFLOW%, lflow%, Pres%, tempPres%
Global REGPOS As Long
Global lastRegPos As Long
Global CREGPOS As Long ' really only integer, but safer this way
Global HREGPOS As Long
Global lfcpos As Integer
' this is 1 for most cases, and 2 for high pressure regulator
Global current_regulator As Integer
Global SlurryPumpSpeedPOS As Long '6.71.123.10
Global slurry_wash_pump_max_flow_cc As Long '6.71.123.10
Global slurry_tube_almost_empty_counts As Long '6.71.123.11
Global slurry_tube_almost_empty_cm As Long '6.71.123.11
Global slurry_tube_almost_full_counts As Long '6.71.123.11
Global slurry_tube_almost_full_cm As Long '6.71.123.11
'Global slurry_tube_csecarea_cm2 As Long '6.71.123.11

Global Slurry_wash_valve As Integer
Global Slurry_tube_vent_valve As Integer
Global Slurry_wash_pump As Integer
Global Slurry_tank_paddle As Integer
Global Slurry_tube_top_shut_off As Integer
Global Slurry_tube_fill_valve As Integer

' these hold the regulator lookup table information
Global reg_table_pos() As Long, reg_table_pres!(), reg_table_size%(1)
Global reg_table_pos2() As Long, reg_table_pres2!()
Global creg_table_pos() As Long, creg_table_pres!(), creg_table_size%
Global regnum As Integer
Global compression_pressure As Double
Global sample_compression_pressure As Double
Global sample_compression_diameter As Double
Global use_sample_compression As Boolean ' true if we use the above two values
Global cyclic_compression_pressure As Single
Global cyclic_compression_timedown As Single
Global cyclic_compression_timeup As Single
Global cyclic_compression_numcycles As Integer

Rem  These are the new global variables for the new CV linearization routines
Global Const MAX_FLOWMETERS = 1000
Global Const MAX_LOHM_DATAPOINTS = 5000

'CV_flow has the lohm calib. flows, a set of points for each flowmeter (fm, flow)
'intermediate_CV_Value has the lohm resistance values, a set for each flowmeter and press. gauge (fm, pg, lohm)
Global QVol, CV_flow!(MAX_FLOWMETERS, MAX_LOHM_DATAPOINTS), intermediate_CV_Value!(MAX_FLOWMETERS, 1, MAX_LOHM_DATAPOINTS), cv_table_size%(MAX_FLOWMETERS)
Global First_Good_CV_Index%(1) ' 0 for high pressure gauge, 1 for low pressure gauge
Global lohmStartMultiplier As Single                     ' User-selectable starting flow value for lohm calculation
Global current_lohm_path$                               ' String to hold name of current lohm table

Rem added by jsd jan 2001
Global v2now_open As Boolean
Global Done_with_v2 As Boolean
Global cvpoints As Integer
Global Disable_CV As Boolean
Global cv_withmulti_v2 As Boolean
Global MedFM_CV_Disable As Boolean
Global Cv_reg_inc As Integer
Global bad_cv_correction As Boolean
Rem god I love global variables
Global Lohm_Ratio As Single
Global cv_flag As Boolean
Global cv_warning_flag As Boolean

' the following are testing parameters set by user for current test
Global OutFilename$(), OutLogFileName$(), tfactor As Single, sid$(), Line1$(), Line2$()
Global surfTen() As Single, fluid$(), operator$(), lot_number$()
Global TPFWET$(), TPFDRY$(), Diam() As Single, cyl_len() As Single
Global innerDiam() As Single, outerDiam() As Single
Global minp_set() As Single, maxp_set() As Single
Global thick() As Single, Hold_Press() As Single, Hold_Time() As Single
Global mf_press() As Single, mf_time() As Single
Global Hold_Delay() As Single, Hold_Rate() As Single
Global Step_Time() As Single, TType%(), TMode%(), Gas$, Liquid$(), GasID$, LiquidID$()    ' IDs are for use with other languages
Global path(1) As String, diffpgflow() As Boolean, stop_at_bp As Boolean
Global threestagetest() As Boolean
Global nowait_for_du As Boolean
Global linear_type% ' 0 = linear, 1 = darcy, 2 = sqrt
Global use_fluid_sensor As Boolean, use_temperature As Boolean
Global use_time As Boolean
Global extend_num%, extend_name$(50), extend_value$(50)
'Global advanced_low_flow As Boolean ' extra low flow checked
Global SquarePores As Integer ' if they want to use sqaure pores
Global current_unit% ' current unit being run - usually 1
Global use_min_pressure_in_dry As Boolean
Global runAsPassFail() As Boolean           ' Run as a pass/fail test
Global stopTestOnFail() As Boolean          ' Stop the test if it fails the criteria
Global minPassDiameter() As Single, maxPassDiameter() As Single     ' min and max pass/fail criteria
Global minMedianPass() As Single, maxMedianPass() As Single         'min and max pass/fail for median pore size
Global passFailType() As Integer
Global failed As Boolean

Global runFrazierAsPassFail() As Boolean
Global minFrazierPass() As Single, maxFrazierPass() As Single

' The following are variables for a sample ejection system
Global hasSampleEject As Boolean
Global sampleEjectValve As Integer

' the following are misc. globals
Global Curr_U$ ' current user name
Global manrunning As Boolean ' true if manual control is currently running
Global FrazierRunning As Boolean ' true if we are using the frazier pressure gauge
Global Cancel_Aborted As Single ' test has been canceled
Global Curve_Ave As Integer ' stores if they want curve fit or averaging
Global Aborted As Boolean ' test has been aborted
Global Get_First As Integer ' true if this is first get of capstuff
Global EXE_Path$ ' path of executable
Global T_Select$ ' type of selection box to show
Global IFile$ ' inifile location
Global CSFile$ ' capstuff file location
Global UACFile$ ' User Access Control file location
Global Decimal_Point$ ' character used on users system for decimal point
Global HelpFile$ ' points to help file
Global reply% ' reply from last message box, sometimes needs to be global
Global RunTimer As Single ' when test was started for progress form
Global Got_Text As String ' return text value of getvalue form
Global Got_Value As Single ' return value from getvalue from
Global Got_Value_Check As Integer
Global HKey$ ' key press returned from form - abort
Global HKey2$ ' key press returned from form - turnaround
Global file$ ' current file being tested
Global PWFACTR As Single ' multiplier of valve 2 increment
Global intest As Boolean ' if we are in a test or not
Global RUNNING As Boolean ' if test is running
Global FlowFlag As Boolean ' high flows have been seen
Global V2POS As Long ' valve 2 target position
Global x5 As Single ' return value of gauge read
Global x4 As Long ' count value of gauge read
Global X1FIRST As Long ' count value first read
Global real_atm As Single ' real atmospheric pressure
Global atm_x4(3) As Long ' count value at atmospheric pressure
Global pending$ ' pending commands in manual control
'6.71.123.02 Global command_issued%(74) ' manual control command hit
Global command_issued%(105) ' manual control command hit            '6.72.017
Global want_to_quit_manual_control As Boolean ' self explanatory
Global auto_index% ' index of scan in manual control
Global auto_mode% ' type of auto scan being done
Global v2_plunger_left% ' position of left v2
Global bubblerMV_plunger_left%
Global last_reg_status As Boolean ' true if regulator was just zeroed
Global GaugeCali_Aborted As Boolean
Global TestScreenVisible As Boolean
Global want_to_hold As Boolean, holding As Boolean
Global temperature1() As Single ' stores the first temperature reading
Global temperature2() As Single ' stores the second temperature reading
Global temperature_array_size As Integer
Global nowait_at_beginning As Boolean
Global nowait_at_end As Boolean
Global unitnumber As Integer ' instrument unit number for multiple instruments
Global density() As Single ' These are for WESA
Global mass() As Single
Global dens_unit As String
Global mass_unit As String
Global mass_unit_c As Single
Global dens_unit_c As Single
Global SCDiam As Single ' the diameter of the sample (>diam of o-ring)
Global TargetPercPorosity As Single 'Target Percent Porosity: Used to calculate target thickness '6.71.123.14
Global target_thickness As Single 'Target Thickness: Calculated '6.71.123.14
Global BuildCakeByPressureTargetPressure As Single ''6.71.123.14
Global BuildCakeByPressureEndFlow As Single ''6.71.124.00
Global BuildCakeByFlowTargetFlow As Single '6.71.123.14
Global BuildCakeByFlowEndPressure As Single '6.71.123.14
Global SlurryTubeWashCycleTargetFlow As Single '6.71.123.14
'Global SlurryTankVolume_cc As Single '6.71.123.14
Global WashTankVolume_cc As Single '6.71.123.14
Global BuildCakeByPressureAbort As Boolean '6.71.123.25
Global temperatureControl As Boolean 'AJB 10-22-09

Global Show_Result As Boolean ' if true then calculate and show results at end of test
Global FResult$ ' stores the result for output at end of test
Global user_keypress As Integer ' stored keyascii value of user keypress during test
Global save_setup_data_flag As Boolean ' true if the test setup is to save the data
Global manual_data_path$ ' data path for data logging in liquid vapor manual control
Global manual_data_logging As Boolean ' true if they are doing data logging
Global abort_lv_goto As Boolean
Global uselog As Boolean
Global permeabilityLogging As Boolean
Global logpath As String
Global permeabilityLoggingFile As String
Global leak_test_delay As Single
Global stability_debug As Boolean
Global advanced_settings As Boolean         ' "Use advanced settings only"
'Global regulator_settings As Boolean        ' "use second regulator only"
Global use_second_regulator_only As Boolean

Global start_caprep As Boolean              ' start CapRep at the end of a test
Global minmaxunits As String                ' "p"=start and stop test based on pressure values
                                            ' "d"=start and stop test based on pore diameter
Global pressHoldUnit As String              ' for pressure hold test, PSI/sec or PSI/min?
Global num_PH_AvePoints As Integer          ' Number of points used in press. hold test averaging
Global PH_reading_freq As Single            ' wait time between readings in press. hold test
Global PH_fail_method$                      ' Whether failure is based on dp/dt or just abs(dp)
Global PH_stopOnFail As Boolean             ' Does test stop if it fails?
Global PH_autoscale As Boolean              ' Autoscale y-axis?
Global PH_minY As Single                    ' min. y value if not autoscaling
Global PH_maxY As Single                    ' max. y value if not autoscaling
Global PH_regression As Boolean             ' True if using lin. regression to calculate pressure decay rate
Global LP_mintime As Single                 ' Minimum time (in seconds) to wait between readings during
                                            ' the liquid perm test
Global LP_FlushBeforeTest As Boolean
Global LP_CCsToFlush As Integer
Global LP_FlushPressure As Single

Global LP_DrainAfterTest As Boolean
Global LP_DrainTime As Integer
                                            
Global auto_report_type                     ' Method of showing results at end of test (replaces start_caprep
                                            ' and Show_Result)
Global gasflowconversionfactor As Single    ' for alternative gasses
Global microflowregulator As Boolean        ' true if we correct for back pressure during microflow test
Global norefill As Boolean                  ' true if we don't refill during elevated pressure liquid permeability
Global delaycompressionliquid As Boolean    ' true if we delay compression until after the initial liquid fill
Global MF_linearSeal As Boolean             ' true if user is calculating microflow through a linear seal rather than whole sample
Global MF_sealDiam As Single                ' length of the "linear seal" microflow option
Global MF_innerDiam As Single               ' Inner diameter for a µflow linear seal
Global MF_outerDiam As Single               ' Outer diameter for a µflow linear seal
Global MF_Settle As Boolean                 ' true if we wait for pressure to settle before starting microflow test
Global MF_Settle_pressure As Single         ' pressure fluctuation allowed during settling check
Global MF_Settle_time As Single             ' time during which pressure must not fluctuate too much
Global MF_recordTemperature As Boolean      ' Enable recording of temperature readings to main microflow data file
Global MF_Total_Settling_time As Single     ' time it actually took to settle
Global PS_usingList As Boolean              ' Flag for using a pressure step list
Global PS_path$                             ' Pathname of the test's pressure step list
Global debugMenuVisible As Boolean          ' iff debug menu in TitleScrn is visible to user
' Variables for a single-point gas perm test
Global GP_singlePointTest As Boolean        ' Flag for doing a single point GP test with averaging
Global GP_target As Single                  ' Target pressure in PSI
Global GP_delay As Integer                  ' Delay time at start of test in seconds
Global GP_duration As Integer               ' Test duration in seconds
Global GP_interval As Integer               ' Interval between data points in seconds
Global GP_numavg As Integer                 ' Number of readings averaged together as a single data point
' Variables for Multi Set gas perm test edc 09-19-07
'Global GP_multisetTest As Boolean           'flag signifying a multiSet test gas perm test
'Global GPM_targetVol As Double              'Target volosity chosen by the user in l/s
'Global GPM_numberSets As Integer            'Number of sets in the test
'Global GPM_setDuration As Integer           'Duration of each set
'Global GPM_numberDataPts As Integer         'Number of data points in each set
'Global setsToDo As Integer                  'counter for number if sets in the MultiSet Gas Perm test

Global BP_AutoDetectMethod As Integer
Global BP_MaxFPT As Single
Global BP_MaxDeltaFPT As Single
Global BP_X_Max As Single
Global BP_Y_Max As Single
Global BP_Points As Long
Global BP_UsePressureVsTime As Boolean       ' If it is true, it shows Pressure vs Time graph otherwise
                                             ' it show F/PT graph. We use this variable when we try to find
                                             ' bubble point pressure with cartridge test (the cartridge has large area,
                                             ' which makes use of F/PT questionable.
                    

Global CF_SampleType As Integer             '0 for thru-plane, 1 for in-plane

Rem new globals to make the program hardware independent
Global DAC_under As Long ' underflow count value
Global DAC_zero As Long ' zero volt count value
Global DAC_two As Long ' 2 volt count value
Global DAC_span As Long ' DAC_two - DAC_zero - the span of the converter
Global DAC_over As Long ' overflow count value

Rem new variables for version 7 hardware
Global xignore As Byte ' number of readings to ignore before taking actual readings - 0=0
Global xmult As Byte ' number of multiple readings to take and then average - 0=256
Global xjiffy As Byte 'time value to use for pulsing motor valves
Global readings_counter As Integer
Global numhangs As Long
Global log_comm As Boolean ' if we log communications errors
'Global log_raw As Boolean ' if we log raw pressure values for each data point

Global compregcal As Boolean
Global auto_advanced As Boolean
Global auto_increment As Boolean
Global AutoSamplID As Boolean   'auto increments sample IDs
Global reg_zero_time As Single
'6.71.123.15 Global qcshow(20) As Byte
Global qcshow(27) As Byte
Global linear_unit_index%, thick_unit_index%, mass_unit_index%, dens_unit_index%, press_unit_index%
Global max_bp_pres_dif As Single
Global reverse_flow_controller As Boolean
Global bottom_fill_point As Single ' cm value for penetrometer for initial fill.  defaults to 0 which disables its use.
Global sample_zero_point As Single ' cm value for true zero height of penetrometer.
Global last_penetrometer_reading As Single ' last reading in manual control - used by set buttons
Global penetrometer_start_test_point As Single ' cm value for start of test.  defaults to pen500 value.
Global max_fill_point As Single ' cm value for stop of filling.  defaults to pen500 value
Global minimum_liquid_test_stop_point As Single ' cm value for minimum point where test can be safely stopped.  defaults to 50% of penetrometer.
Global Compression_Increase_Factor As Single
Global dual_stage_compression As Boolean
Global pretreat_time As Single
Global pretreat_flow As Single
Global first_flow_starting_point_percent As Single ' defaults to 100%, which means use SHFP


' Rem statement directly follows inserted code
' **********
' Begin Mettler Balance INI file code entry for use of a balance entered by
' search for Tim Richards on Wednesday 04 05 26
'
Global g_bBalanceNotPenet As Boolean                ' uses H2OPERM variable set to 'B' and Featyre 16 to use with LEP. Search for Tim Richards. 04 05 14
Global g_iMettler_fluid_min As Single               ' (grams) Use in ReadBalanceNotPenet in place of PEN500, PEN20500. Search for Tim Richards. 04 05 26
Global g_iMettler_fluid_max As Single               ' (grams) Use in ReadBalanceNotPenet in place of PEN500, PEN20500. Search for Tim Richards. 04 05 26
Global g_iMettler_fluid_density As Single           ' (grams / mL) Use in ReadBalanceNotPenet in place of PEN500, PEN20500. Search for Tim Richards. 04 05 26
Global g_iMettler_Negative_Counts_Offset As Single  ' put an offset in there so that the AutoTest procedure gets the inital pen counts in range. TAR 040614
Global g_iMettler_MaxFlowSetPoint As Single         ' "Reached 10mL of flow. Stopping test"---ts$ 493. Put a set point in there, not hardcoded. TAR 040728
Global g_iBalanceNotPenet_SettlingTime As Single    ' "Enter settling time once target pressure is attained:") 'TAR 040804
Global g_bBalanceNotPenet_ZeroPoint As Single       ' Subtract off the actual reading from the balance
'
' End Mettler Balance INI file code entry by Tim Richards 04 05 26
' **********

' new variable to skip data points in dry curve until you get to a certain flow rate
Global min_flow_in_dry As Single ' flow rate, in cc/min
Global use_min_flow_in_dry As Boolean ' if this is turned on or not

' **********
' Begin code inserted by search for Tim Richards on Wednesday 6/30/04
' use records to assess the readings coming in from the balance and the regulator in Run_Elev_LqPerm
'
'Type MeasurementsRecord
'    Pressure As Single
'    mass As Single
'    time As Single
'    diffT2time As Single
'End Type

'Type FinalMeasurements
'    Volume As Single
'    Flow As Single
'    difftime As Single
'    MeanPressure As Single
'    MeanPoints As Integer
'    MeasurementPoints As Integer
'End Type
'
' End code inserted by Tim Richards 6/30/04
' **********


Rem new variables (array) for creation of separate file showing raw and corrected pressures
Rem added JSD April 22, 2001
Global Lohm_P() As Single
Global Lohm_Val() As Single
Global Lohm_F() As Single
Global Raw_P() As Single
Global lohm_counter As Single
Global size_of_lohm_array As Single

'JF 10-18-2010 Adding variable to determine if Lohm Calibration is running
Global runningLohmCalibration As Boolean

Global curve_perc As Single
Global curve_nump As Integer
Global curve_maxd As Single

Global preloaded_sample As Boolean
Global first_test_setup As Boolean ' used by those who call testscrn
Global simpleqc_enable As Boolean

' the following are new with the cftfile routines
' return value for the glsel routine - replaces old global "Air$"
Global selected_gas_or_liquid As tpgl
Global flow_cal(40) As Single

' the following store requested changes to the gauge range variables
' they are used to delay the actual changes until between when gauges
' are being read to avoid incorrect readings when you change ranges
' This is used in manual control only.
' When one of these is set to -1, that means that there is no requested change
Global future_pres%, future_lflow%, future_hflow%, future_DPress%, future_CPress%, future_EPress%

Global maxLowPressureGaugePressure As Single

' replaced menu selection of chambers with global array and new form
Global selchamber(10) As Boolean

' true if we are doing special integrity porometry (or gas permeability)
Global integrity_porometry As Boolean

'For holding parameters
Global BPParams As String
Global DRYParams As String
Global WETParams As String

' For looping demo tests
Global LoopingDemo As Boolean
Global LDtempValue As Single               ' temp value for saving data

' For debugging
Global debugH20Perm As Boolean             ' write to debug log during liquid perm.
Global debugBP As Boolean                  ' write to debug log during bubble point

' For debugging functions
Global debugRunCPass As Boolean

' for remembering overshoot in deltap=0 elevated pressure liquid permeability
Global lperm_last_target_pressure As Single
Global lperm_last_pressure_overshoot As Single

' 6.71.67 these are now replaced by DryChamberTemperature, WetChamberTemperature, and RecirculationTemperature
'' for external watlow temperature controller
'Global watlow_com_number As Integer ' defaults to 0, meaning no external watlow
'Global watlow_last_temperature As Single ' remember last temperature in case the watlow fails to respond
'Global using_watlow As Boolean ' true if we are using the Watlow during this test
Global watlow_last_temperature(2) As Single ' remember last temperature for each of 2 channels in case the watlow fails to respond

' for new valve 23
Global valve_23_exists As Boolean

' 6.71.20 begin
' for reading both high flow meters during the test and swapping regulators
Global suspend_v10 As Boolean ' true if v10 is not supposed to move
Global switch_high_flow_enabled As Boolean ' true if we allow switching of high flow meter
Global using_hflow1 As Boolean ' true if we are currently using hflow1 but the test can eventually use hflow2
Global hflow1_max_index As Integer ' initially 0, the number of secondary flow data points that have been taken so far
Global second_regulator_starting_point As Integer ' defaults to 0, otherwise is the point on the second regulator that equals the pressure of 4000 counts on the first regulator
Global using_low_regulator As Boolean ' true if we are using the lower regulator but want to switch at some point if we have to
' 6.71.20 end

' For automatic curve fitting when a test is done
Global autoCurveFit As Boolean

' For temperature control at end of a test
Global zeroTempAtEndOfTest As Boolean

' Bubble Point Time Log
Global BPTLEnable As Boolean
Global BPTLInterval As Long ' in 0.1 second increments
Global BPTLMaxPoints As Long
Global BPTLNumPoints As Long
Global BP_PointDetectionCount As Integer
Global BPTestStopTimeInterval As Integer ' time in seconds to stop the BP test if pressure does increase (leakage)
Global bubWaitTime As Integer
Global bubPressOnWait As Boolean

' For variable delay time when closing piston -- gives the hardware a chance to seal for slower pistons
' 6.71.61
Global pistonDelayTime As Integer

' 6.71.64 - for running multiple gasperm tests and averaging the results together
Global GP_multiAverageTest As Boolean
Global gP_numAvgTests As Integer            ' Number of tests to run when in averaging mode
Global GP_multiAvgCounter As Integer        ' Counter to keep track of tests run

' 6.71.67
' 0 = no temperature probe at this location
' 1 or 2 = use V6 or V7 aux channel 27 or 30 (configuration "G1" or "G2")
' 3 = use Rabbit pass-through serial port 1, watlow channel A (configuration "R1A")
' 4 = use Rabbit pass-through serial port 1, watlow channel B (configuration "R1B")
' 5 = use Rabbit pass-through serial port 2, watlow channel A (configuration "R2A")
' etc.
' -1 = use external PC serial port 1, watlow channel A (configuration "C1A")
' -2 = use external PC serial port 1, watlow channel B (configuration "C1B")
' -3 = use external PC serial port 2, watlow channel A (configuration "C1A")
' etc.
' Note that if you use two external PC serial port probes, they must both be on the same port (but different
'  watlow channels).  You can't currently use three external PC serial port probes.
Global dryChamberTemperature As Integer, wetChamberTemperature As Integer, reservoirTemperature As Integer
Global cabinetTemperature As Integer, airTemperature As Integer, bubblerTemperature As Integer
Global hydroHeadTemperature As Integer, mullenTemperature As Integer

Global watlowViaModbus(3) As Boolean
Global athena As Integer ' normally 0.  Set to 2 if using 2 channel Athena controller instead of Watlow

'10-19-05 next globals added by Edward Corvinelli for the selection of gas or liquids from one form GLSel.frm
'and the access of this information from a veriety of forms.
' 3-13-06 modified by rvw
Global NeedLiquid As Boolean
Global NeedGas As Boolean
Global NeedFluid As Boolean

' 3-16-06 rvw
Global FrazierChamberValve As Integer
Global FrazierPiston As Boolean

' 5-4-06 rvw
Global network_connection_enabled As Boolean
Global network_connected As Boolean

' 9-15-06 rvw
Global penet_refill_delay As Single

'12-06-06 edc
Global lngBorderColor As Long      'will be a shared value for all forms borders
Global SubCaption As String     'will share a string that has part of the EXE_Path and the app name

' 6.71.102
Global air_inlets As Integer
Global current_air_inlet As Integer
Global air_inlet_1_max_p As Single

' 03-08-07 edc
'Global MinimumReliableHighFlowRate As Double

Global pen_max_counts As Long
Global pen_min_counts As Long
Global pen_span As Long

Global leakTestPassPercent As Single

' 11-20-07
Global minbubflow As Single

'10-20-09 AJB
Global temperatureSetPoint As Single

' 12-30.2009 by JF
Global pneumaticSwitchValveForPiston As Boolean

' 1-13-2010 by JF
Global readAllFlows As Boolean

'1-21-2010 by JF
Global bpRunMultipleTests As Boolean
Global bpTestCount As Integer

'8-22-2010 by JF
Global lpRunMultipleTests As Boolean
Global lpTestCount As Integer

'2-9-2010 by JF
Global useAdditionalInfo As Boolean
Global numberOfAdditionalInfoLines As Integer
Global infoLineHeaders() As String
Global infoLineValues() As String

'2-10-2010 by JF
Global ReserveTankLevelChannel As Integer
Global ReserveTankLevelZero As Long
Global ReserveTankLevelSpan As Long
'Global ReserveTankFillLight As Boolean
Global ReserveTankFillLightValve As Integer
Global ReserveTankRefillPercent As Integer
Global ReserveTankLevelMin As Integer

' 5-22-12 added for dual chamber rotating sample bubble point tester
Global ChamberLiquidLevelChannel(2) As Integer
Global ChamberLiquidLevelZero(2) As Long
Global ChamberLiquidLevelSpan(2) As Long
' 6-18-12 added for dual chamber to control liquid level during the fill and drain
Global ChamberLiquidLevelMin(2)  As Single
Global ChamberLiquidLevelMax(2)  As Single

'2-10-2010 by JF
'Added for automatic test temperature control
Global dryChamberTargetTemperature() As Single
Global wetChamberTargetTemperature() As Single
Global reservoirTargetTemperature() As Single
Global airTargetTemperature() As Single
Global bubblerTargetTemperature() As Single
Global cabinetTargetTemperature() As Single
Global hydroHeadTargetTemperature() As Single
Global mullenTargetTemperature() As Single
Global useTemperatureControlForAuto() As Boolean
Global setTemperatureForAuto() As Boolean
Global delayTestForTemperature() As Boolean
'Global delayTestForTemperatureControl() As Boolean
Global minimumPossibleTemperature As Single
Global maximumPossibleTemperature As Single

'2-15-10
'Added for humidity controls in the Baxter machine
'Type used for holding humidty table data
Public Type humidityTableData
    bubblerMVPos As Single
    humidity As Single
    flowRate As Single
End Type

Global hasHumidityControls As Boolean
Global humidityGaugeNumber As Integer
Global humidityRegulatorPosition As Integer

' variables to control humidity, adjustable from preferences screen
Global enableHumidityControlForAutoTests 'flag to say whether or not we're calling goToTargetHumidity
Global recordHumidityForAutoTests As Boolean 'flag to say whether or not we're recording data from the humidity sensor
Global targetHumidity As Single
Global humidityFile As Boolean 'flag to indicate if we're reading / writing a file with humidity data
Global goToHumidityMaxWaitTime As Single 'seconds
Global goToHumidityMinWaitTime As Single 'seconds
Global goToHumidityTolerance As Single 'percent
Global stableHumidityMaxWaitTime As Single 'seconds
Global stableHumidityMinWaitTime As Single 'seconds
Global stableHumidityTolerance As Single 'percent
Global stableHumiditySleepTime As Long 'milliseconds
Global initialHumidityWaitTime As Single 'seconds
Global minHumidityAdjustmentFlow As Single 'cc's, min flow to start adjusting humidity, or 0 if don't want to use

Global lastGoodHumidity As Single 'store last saved humidity value that should be within tolerance of target

Global humidityTable() As humidityTableData
Global humidityTableMaxSize As Integer
Global humidityTableCurrentSize As Integer

Global zeroRegVentTime As Long 'number of milliseconds to vent the pressure above the bubbler when zeroing the regulator
                               'in machines that control humidity

Global goToHumidityCounter As Integer 'counter for how many times we've called goToTargetHumidity()
Global goToHumidityWaitTimeCounter As Long 'hold the amount of time we waited to reach a stable humidity for the last time we called goToTargetHumidity

Global afterBubblePoint As Boolean

'2-23-10
'more humidity sensor stuff
Global humiditySensorZeroCounts As Integer
Global humiditySensorFullCounts As Long
Global tankFullCounts As Long
Global tankZeroCounts As Integer
Global HumidityValues() As Single

'3-11-10
'Adding a variable to allow a porometry test to skip motor valve error messages
Global ignoreMVErrors As Boolean

'12-2-10
'Adding variables for Cornings sample chamber diverter valve
Global sampleChamberDiverterValve As Integer
Global divertSampleChamber As Boolean

'2-8-11
Global useNewSinglePointRoutine As Boolean

'3-17-2011
Global showEndPassCause As Boolean

'6-27-2011
Global calibrationComplete As Boolean

'7-18-2011
Global switchingMVs As Boolean

'07-12-12
'Valve number which turns on and off Pump to fill cartridge tester chambers
Global PumpValveNumber As Integer

' 02-23-15
' Purge routine related parameters
Global showPurgeOption As Boolean
Global purgeStopCounts As Long

' The time when we started tracking time
Global TrackTime As Long

' Remap options - Mix
Global Remap As String
Global UseRemap As Boolean
Global ShowRemapOption As Boolean
Global RemapOptionCaption As String
Global ValveRemap(36) As Integer
Global GaugeRemap(42) As Integer
Global FireOnRemap As String
Global FireOnUnRemap As String

' Pretest valve flips - Mix
Global Flip1Show As Boolean
Global Flip2Show As Boolean
Global Flip1Name As String
Global Flip2Name As String
Global Flip1Raw As String
Global Flip2Raw As String
Global Flip1Valve As Integer
Global Flip2Valve As Integer

' Use regulator calibration for LP test
Global UseRegCalForLP As Boolean

' Athenas
Global UseAthena1 As Boolean
Global Athena1Units As String
Global Athena1Channel As Integer
Global Athena1Target As Single
Global Athena1InTest As Boolean
Global Athena1NeverInTest As Boolean
Global UseAthena2 As Boolean
Global Athena2Units As String
Global Athena2Channel As Integer
Global Athena2Target As Single
Global Athena2InTest As Boolean
Global Athena2NeverInTest As Boolean

' Reduce flow rate option
Global ReduceFlowAtTarget As Boolean
Global ReduceFlowPressureTarget As Single
Global NeedToWatchPressForReduction As Boolean
Global oldBloop&

' Re-range P2 gauge for bubble point tests
Global NeedP20 As Boolean
