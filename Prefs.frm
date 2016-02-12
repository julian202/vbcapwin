VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form prefsForm 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   8670
   ClientLeft      =   4965
   ClientTop       =   2595
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame13 
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   7440
         TabIndex        =   2
         Top             =   7920
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5520
         TabIndex        =   1
         Top             =   7920
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7575
         Left            =   240
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   13361
         _Version        =   393216
         Tabs            =   15
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "Prefs.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label17"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "setFontButton"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame5"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame4"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Combo1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Tests"
         TabPicture(1)   =   "Prefs.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "depressurizeCheck"
         Tab(1).Control(1)=   "regulatorCheck"
         Tab(1).Control(2)=   "testGasCheck"
         Tab(1).Control(3)=   "AutoSamplIDcheck"
         Tab(1).Control(4)=   "Frame2"
         Tab(1).Control(5)=   "autoincCheck"
         Tab(1).Control(6)=   "hidePromptsCheck"
         Tab(1).Control(7)=   "minPressCheck"
         Tab(1).Control(8)=   "advancedCheck"
         Tab(1).Control(9)=   "Frame9"
         Tab(1).Control(10)=   "curveFitCheck"
         Tab(1).Control(11)=   "minFlowText"
         Tab(1).Control(12)=   "minFlowCheck"
         Tab(1).Control(13)=   "Label13"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "Pressure Hold"
         TabPicture(2)   =   "Prefs.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "pLabel(3)"
         Tab(2).Control(1)=   "pLabel(2)"
         Tab(2).Control(2)=   "pLabel(1)"
         Tab(2).Control(3)=   "pLabel(0)"
         Tab(2).Control(4)=   "startingPText"
         Tab(2).Control(5)=   "and1"
         Tab(2).Control(6)=   "and2"
         Tab(2).Control(7)=   "minEndPsiText"
         Tab(2).Control(8)=   "Frame7"
         Tab(2).Control(9)=   "testLength"
         Tab(2).Control(10)=   "delayTime"
         Tab(2).Control(11)=   "Frame3"
         Tab(2).Control(12)=   "numAvePoints"
         Tab(2).Control(13)=   "pholdFreq"
         Tab(2).Control(14)=   "startingP1"
         Tab(2).Control(15)=   "startingP2"
         Tab(2).Control(16)=   "endingP1"
         Tab(2).Control(17)=   "endingP2"
         Tab(2).ControlCount=   18
         TabCaption(3)   =   "Liq. Perm."
         TabPicture(3)   =   "Prefs.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "pLabel(6)"
         Tab(3).Control(1)=   "lblLPFlushCCs"
         Tab(3).Control(2)=   "lblLPFlushPressure"
         Tab(3).Control(3)=   "pLabel(9)"
         Tab(3).Control(4)=   "lblV6LPRegIncAmount"
         Tab(3).Control(5)=   "lblV6LPRegIncWait"
         Tab(3).Control(6)=   "delaycompressionliquidcheck"
         Tab(3).Control(7)=   "norefillcheck"
         Tab(3).Control(8)=   "lpMintime"
         Tab(3).Control(9)=   "chkLPFlushBeforeTest"
         Tab(3).Control(10)=   "txtLPFlushCCs"
         Tab(3).Control(11)=   "txtLPFlushPressure"
         Tab(3).Control(12)=   "chkLPDrainAfterTest"
         Tab(3).Control(13)=   "txtLpDrainTime"
         Tab(3).Control(14)=   "txtV6LPRegIncAmount"
         Tab(3).Control(15)=   "txtV6LPRegIncWait"
         Tab(3).ControlCount=   16
         TabCaption(4)   =   "Calibration"
         TabPicture(4)   =   "Prefs.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lohmCaption2"
         Tab(4).Control(1)=   "lohmCaption"
         Tab(4).Control(2)=   "Label7"
         Tab(4).Control(3)=   "Label12"
         Tab(4).Control(4)=   "lohmCaption3"
         Tab(4).Control(5)=   "lohmCaption4"
         Tab(4).Control(6)=   "lohmCaption5"
         Tab(4).Control(7)=   "lohmTimeoutText"
         Tab(4).Control(8)=   "lohmPercent"
         Tab(4).Control(9)=   "lohmFlowText"
         Tab(4).Control(10)=   "lohmRegulatorText"
         Tab(4).Control(11)=   "lohmToleranceText"
         Tab(4).ControlCount=   12
         TabCaption(5)   =   "Microflow"
         TabPicture(5)   =   "Prefs.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "mf_settle_label2"
         Tab(5).Control(1)=   "mf_settle_label1"
         Tab(5).Control(2)=   "mf_temperature_check"
         Tab(5).Control(3)=   "mf_settle_time_text"
         Tab(5).Control(4)=   "mf_settle_pressure_text"
         Tab(5).Control(5)=   "mf_settle_check"
         Tab(5).Control(6)=   "linSealCheck"
         Tab(5).Control(7)=   "microflowregulatorcheck"
         Tab(5).ControlCount=   8
         TabCaption(6)   =   "Curve Fit"
         TabPicture(6)   =   "Prefs.frx":00A8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "cflabel(2)"
         Tab(6).Control(1)=   "cflabel(1)"
         Tab(6).Control(2)=   "cflabel(0)"
         Tab(6).Control(3)=   "Label10"
         Tab(6).Control(4)=   "cfMaxPSI"
         Tab(6).Control(5)=   "cfPercentError"
         Tab(6).Control(6)=   "cfNumPoints"
         Tab(6).ControlCount=   7
         TabCaption(7)   =   "Gas Perm."
         TabPicture(7)   =   "Prefs.frx":00C4
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Label11(4)"
         Tab(7).Control(1)=   "Frame12"
         Tab(7).Control(2)=   "Frame11"
         Tab(7).Control(3)=   "Frame14"
         Tab(7).ControlCount=   4
         TabCaption(8)   =   "Special Options"
         TabPicture(8)   =   "Prefs.frx":00E0
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Line2"
         Tab(8).Control(1)=   "BPTLMaxLabel"
         Tab(8).Control(2)=   "BPTLSecondsLabel"
         Tab(8).Control(3)=   "Line1"
         Tab(8).Control(4)=   "Label14"
         Tab(8).Control(5)=   "Label25"
         Tab(8).Control(6)=   "BubblerCheck"
         Tab(8).Control(7)=   "BPTLMaxText"
         Tab(8).Control(8)=   "BPTLSecondsText"
         Tab(8).Control(9)=   "BPTLCheck"
         Tab(8).Control(10)=   "zeroTempCheck"
         Tab(8).Control(11)=   "Text6"
         Tab(8).ControlCount=   12
         TabCaption(9)   =   "AutoFile Options"
         TabPicture(9)   =   "Prefs.frx":00FC
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Frame15"
         Tab(9).Control(1)=   "Frame16"
         Tab(9).Control(2)=   "Frame17"
         Tab(9).Control(3)=   "chkEnabled"
         Tab(9).Control(4)=   "chkFolder"
         Tab(9).Control(5)=   "chkSeperate"
         Tab(9).Control(6)=   "Frame18"
         Tab(9).ControlCount=   7
         TabCaption(10)  =   "Auto Wet Feature"
         TabPicture(10)  =   "Prefs.frx":0118
         Tab(10).ControlEnabled=   0   'False
         Tab(10).Control(0)=   "txtPumpSpeed"
         Tab(10).Control(1)=   "time4"
         Tab(10).Control(2)=   "optWetVolume"
         Tab(10).Control(3)=   "txtAutoWetVolume"
         Tab(10).Control(4)=   "optWetTime"
         Tab(10).Control(5)=   "time3"
         Tab(10).Control(6)=   "time2"
         Tab(10).Control(7)=   "time1"
         Tab(10).Control(8)=   "auto_wet_check"
         Tab(10).Control(9)=   "lblPumpSpeedDesc"
         Tab(10).Control(10)=   "lblPumpSpeed"
         Tab(10).Control(11)=   "time4Label"
         Tab(10).Control(12)=   "Label26"
         Tab(10).Control(13)=   "test_unsupport"
         Tab(10).Control(14)=   "Label29"
         Tab(10).Control(15)=   "drain_time"
         Tab(10).Control(16)=   "Label28"
         Tab(10).Control(17)=   "Label27"
         Tab(10).Control(18)=   "soak_time"
         Tab(10).Control(19)=   "wet_time"
         Tab(10).ControlCount=   20
         TabCaption(11)  =   "Additional Info"
         TabPicture(11)  =   "Prefs.frx":0134
         Tab(11).ControlEnabled=   0   'False
         Tab(11).Control(0)=   "lstInfoLines"
         Tab(11).Control(1)=   "txtNumberOfInfoLines"
         Tab(11).Control(2)=   "chkEnableAdditionalInfo"
         Tab(11).Control(3)=   "lblInfoHeaders"
         Tab(11).Control(4)=   "lblNumOfInfoLines"
         Tab(11).ControlCount=   5
         TabCaption(12)  =   "Humidity Options"
         TabPicture(12)  =   "Prefs.frx":0150
         Tab(12).ControlEnabled=   0   'False
         Tab(12).Control(0)=   "chkEnableHumidityControl"
         Tab(12).Control(1)=   "txtMininmumAdjustmentFlow"
         Tab(12).Control(2)=   "txtInitialHumidityWaitTime"
         Tab(12).Control(3)=   "txtStabilitySleepTime"
         Tab(12).Control(4)=   "txtStabilityTolerance"
         Tab(12).Control(5)=   "txtMinStabilityWaitTime"
         Tab(12).Control(6)=   "txtMaxStabilityWaitTime"
         Tab(12).Control(7)=   "txtTargetTolerance"
         Tab(12).Control(8)=   "txtMinAdjustmentWaitTime"
         Tab(12).Control(9)=   "txtMaxAdjustmentWaitTime"
         Tab(12).Control(10)=   "chkRecordHumidityForAutoTests"
         Tab(12).Control(11)=   "txtTargetHumidity"
         Tab(12).Control(12)=   "Label47"
         Tab(12).Control(13)=   "Label37"
         Tab(12).Control(14)=   "Label36"
         Tab(12).Control(15)=   "Label35"
         Tab(12).Control(16)=   "Label34"
         Tab(12).Control(17)=   "Label33"
         Tab(12).Control(18)=   "Label32"
         Tab(12).Control(19)=   "Label31"
         Tab(12).Control(20)=   "Label30"
         Tab(12).Control(21)=   "lblTargetHumidity"
         Tab(12).ControlCount=   22
         TabCaption(13)  =   "Bubble Point"
         TabPicture(13)  =   "Prefs.frx":016C
         Tab(13).ControlEnabled=   0   'False
         Tab(13).Control(0)=   "Command3"
         Tab(13).Control(1)=   "savePreBPdataCheck"
         Tab(13).Control(2)=   "txtBPPointDetectionCount"
         Tab(13).Control(3)=   "pLabel(8)"
         Tab(13).ControlCount=   4
         TabCaption(14)  =   "User Access"
         TabPicture(14)  =   "Prefs.frx":0188
         Tab(14).ControlEnabled=   0   'False
         Tab(14).Control(0)=   "lblUsers"
         Tab(14).Control(1)=   "Label49"
         Tab(14).Control(2)=   "chkEnableUAC"
         Tab(14).Control(3)=   "lstUsers"
         Tab(14).Control(4)=   "cmdAddUser"
         Tab(14).Control(5)=   "cmdRemoveUser"
         Tab(14).Control(6)=   "optAccountType(0)"
         Tab(14).Control(7)=   "optAccountType(1)"
         Tab(14).Control(8)=   "cmdChangePassword"
         Tab(14).ControlCount=   9
         Begin VB.CommandButton Command3 
            Caption         =   "Sample Options"
            Height          =   570
            Left            =   -74235
            TabIndex        =   275
            Top             =   2205
            Width           =   2340
         End
         Begin VB.TextBox endingP2 
            Height          =   285
            Left            =   -71640
            TabIndex        =   272
            Text            =   "500"
            Top             =   2640
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox endingP1 
            Height          =   285
            Left            =   -74520
            TabIndex        =   271
            Text            =   "0"
            Top             =   2640
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox startingP2 
            Height          =   285
            Left            =   -71640
            TabIndex        =   270
            Text            =   "500"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox startingP1 
            Height          =   285
            Left            =   -74520
            TabIndex        =   268
            Text            =   "0"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtV6LPRegIncWait 
            Height          =   285
            Left            =   -74880
            TabIndex        =   265
            Text            =   "Text1"
            Top             =   6000
            Width           =   735
         End
         Begin VB.TextBox txtV6LPRegIncAmount 
            Height          =   285
            Left            =   -74880
            TabIndex        =   263
            Text            =   "Text1"
            Top             =   5640
            Width           =   735
         End
         Begin VB.TextBox txtLpDrainTime 
            Height          =   285
            Left            =   -74520
            TabIndex        =   261
            Text            =   "Text1"
            Top             =   4680
            Width           =   735
         End
         Begin VB.CheckBox chkLPDrainAfterTest 
            Caption         =   "Drain After Elevated Test"
            Height          =   255
            Left            =   -74880
            TabIndex        =   260
            Top             =   4320
            Width           =   4215
         End
         Begin VB.CheckBox savePreBPdataCheck 
            Caption         =   "Save Pre-Bubble Point Data"
            Height          =   255
            Left            =   -74880
            TabIndex        =   259
            Top             =   1560
            Width           =   3495
         End
         Begin VB.CheckBox depressurizeCheck 
            Caption         =   "Depressurize before starting test"
            Height          =   255
            Left            =   -70440
            TabIndex        =   258
            Top             =   3840
            Width           =   3495
         End
         Begin VB.TextBox txtLPFlushPressure 
            Height          =   285
            Left            =   -74520
            TabIndex        =   256
            Text            =   "Text1"
            Top             =   3720
            Width           =   735
         End
         Begin VB.TextBox txtLPFlushCCs 
            Height          =   285
            Left            =   -74520
            TabIndex        =   254
            Text            =   "Text1"
            Top             =   3360
            Width           =   735
         End
         Begin VB.CheckBox chkLPFlushBeforeTest 
            Caption         =   "Flush Before Elevated Test"
            Height          =   255
            Left            =   -74880
            TabIndex        =   253
            Top             =   3000
            Width           =   4215
         End
         Begin VB.CommandButton cmdChangePassword 
            Caption         =   "Change Password"
            Height          =   375
            Left            =   -70440
            TabIndex        =   252
            Top             =   2760
            Width           =   2175
         End
         Begin VB.OptionButton optAccountType 
            Caption         =   "User"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   -69120
            TabIndex        =   251
            Top             =   1920
            Width           =   975
         End
         Begin VB.OptionButton optAccountType 
            Caption         =   "Admin"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   -70320
            TabIndex        =   250
            Top             =   1920
            Width           =   975
         End
         Begin VB.CommandButton cmdRemoveUser 
            Caption         =   "Remove"
            Height          =   375
            Left            =   -73440
            TabIndex        =   248
            Top             =   5520
            Width           =   975
         End
         Begin VB.CommandButton cmdAddUser 
            Caption         =   "Add"
            Height          =   375
            Left            =   -74400
            TabIndex        =   247
            Top             =   5520
            Width           =   975
         End
         Begin VB.ListBox lstUsers 
            Height          =   3570
            Left            =   -74400
            TabIndex        =   245
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CheckBox chkEnableUAC 
            Caption         =   "Enable User Access Control"
            Height          =   255
            Left            =   -74760
            TabIndex        =   244
            Top             =   1200
            Width           =   3255
         End
         Begin VB.TextBox txtPumpSpeed 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -74160
            TabIndex        =   242
            Text            =   "TIME"
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox time4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -73560
            TabIndex        =   239
            Text            =   "TIME"
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtBPPointDetectionCount 
            Height          =   285
            Left            =   -74880
            TabIndex        =   237
            Text            =   "Text1"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox lohmToleranceText 
            Height          =   285
            Left            =   -74880
            TabIndex        =   235
            Text            =   "Text7"
            Top             =   5400
            Width           =   615
         End
         Begin VB.CheckBox regulatorCheck 
            Caption         =   "Use second regulator only"
            Height          =   255
            Left            =   -70440
            TabIndex        =   234
            Top             =   3480
            Width           =   2535
         End
         Begin VB.CheckBox chkEnableHumidityControl 
            Caption         =   "Enable Humidity Control"
            Height          =   255
            Left            =   -74520
            TabIndex        =   233
            Top             =   1560
            Width           =   3555
         End
         Begin VB.TextBox txtMininmumAdjustmentFlow 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   231
            Top             =   6720
            Width           =   735
         End
         Begin VB.TextBox txtInitialHumidityWaitTime 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   223
            Top             =   6240
            Width           =   735
         End
         Begin VB.TextBox txtStabilitySleepTime 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   222
            Top             =   5760
            Width           =   735
         End
         Begin VB.TextBox txtStabilityTolerance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   221
            Top             =   5280
            Width           =   735
         End
         Begin VB.TextBox txtMinStabilityWaitTime 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   220
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox txtMaxStabilityWaitTime 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   218
            Top             =   4320
            Width           =   735
         End
         Begin VB.TextBox txtTargetTolerance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   217
            Top             =   3840
            Width           =   735
         End
         Begin VB.TextBox txtMinAdjustmentWaitTime 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   216
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox txtMaxAdjustmentWaitTime 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   215
            Top             =   2880
            Width           =   735
         End
         Begin VB.OptionButton optWetVolume 
            Height          =   375
            Left            =   -73920
            TabIndex        =   213
            Top             =   2760
            Width           =   255
         End
         Begin VB.TextBox txtAutoWetVolume 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -73560
            TabIndex        =   212
            Text            =   "Volume"
            Top             =   2760
            Width           =   735
         End
         Begin VB.OptionButton optWetTime 
            Height          =   375
            Left            =   -73920
            TabIndex        =   211
            Top             =   1800
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.CheckBox chkRecordHumidityForAutoTests 
            Caption         =   "Record Humidity"
            Height          =   255
            Left            =   -74520
            TabIndex        =   209
            Top             =   1200
            Width           =   3555
         End
         Begin VB.TextBox txtTargetHumidity 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71760
            TabIndex        =   208
            Top             =   2040
            Width           =   735
         End
         Begin VB.ListBox lstInfoLines 
            Height          =   3570
            ItemData        =   "Prefs.frx":01A4
            Left            =   -73920
            List            =   "Prefs.frx":01A6
            TabIndex        =   206
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox txtNumberOfInfoLines 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   -72360
            TabIndex        =   205
            Top             =   1680
            Width           =   735
         End
         Begin VB.CheckBox chkEnableAdditionalInfo 
            Caption         =   "Enable Additional Information"
            Height          =   255
            Left            =   -74640
            TabIndex        =   203
            Top             =   1200
            Width           =   3555
         End
         Begin VB.CheckBox testGasCheck 
            Caption         =   "Test Gas Connection"
            Height          =   255
            Left            =   -70440
            TabIndex        =   202
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   -74280
            TabIndex        =   200
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox time3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -74160
            TabIndex        =   197
            Text            =   "TIME"
            Top             =   4080
            Width           =   735
         End
         Begin VB.TextBox time2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -74160
            TabIndex        =   191
            Text            =   "TIME"
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox time1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -73560
            TabIndex        =   190
            Text            =   "TIME"
            Top             =   1800
            Width           =   735
         End
         Begin VB.CheckBox auto_wet_check 
            Caption         =   "Enable Auto Wet Sample Process"
            Height          =   255
            Left            =   -74640
            TabIndex        =   189
            Top             =   1080
            Width           =   3555
         End
         Begin VB.Frame Frame18 
            Caption         =   "Date Format Settings"
            Height          =   2055
            Left            =   -70200
            TabIndex        =   167
            Top             =   4860
            Width           =   3615
            Begin VB.OptionButton radDate_2 
               Caption         =   "YYYY - Month - DD"
               Height          =   255
               Left            =   240
               TabIndex        =   178
               Top             =   1680
               Width           =   2055
            End
            Begin VB.OptionButton radDate_8 
               Caption         =   "Month - YYYY"
               Height          =   255
               Left            =   1920
               TabIndex        =   175
               Top             =   1320
               Width           =   1575
            End
            Begin VB.OptionButton radDate_7 
               Caption         =   "Month - DD"
               Height          =   255
               Left            =   1920
               TabIndex        =   174
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton radDate_6 
               Caption         =   "Month - YY"
               Height          =   255
               Left            =   1920
               TabIndex        =   173
               Top             =   960
               Width           =   1455
            End
            Begin VB.OptionButton radDate_4 
               Caption         =   "YY - Month - DD"
               Height          =   255
               Left            =   1920
               TabIndex        =   172
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton radDate_5 
               Caption         =   "Month - DD - YY"
               Height          =   255
               Left            =   240
               TabIndex        =   171
               Top             =   1320
               Width           =   1575
            End
            Begin VB.OptionButton radDate_3 
               Caption         =   "Month - DD - YYYY"
               Height          =   255
               Left            =   240
               TabIndex        =   170
               Top             =   960
               Width           =   1695
            End
            Begin VB.OptionButton radDate_1 
               Caption         =   "MM - DD - YYYY"
               Height          =   255
               Left            =   240
               TabIndex        =   169
               Top             =   600
               Width           =   1695
            End
            Begin VB.OptionButton radDate_0 
               Caption         =   "MM - DD - YY"
               Height          =   255
               Left            =   240
               TabIndex        =   168
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.CheckBox chkSeperate 
            Caption         =   "Seperate data in names?"
            Height          =   255
            Left            =   -68880
            TabIndex        =   166
            Top             =   1260
            Width           =   2295
         End
         Begin VB.CheckBox chkFolder 
            Caption         =   "Create Folders?"
            Height          =   255
            Left            =   -70440
            TabIndex        =   165
            Top             =   1260
            Width           =   1815
         End
         Begin VB.CheckBox chkEnabled 
            Caption         =   "Enabled"
            Height          =   255
            Left            =   -71520
            TabIndex        =   164
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Frame Frame17 
            Caption         =   "Filename Generator Settings"
            Height          =   2895
            Left            =   -71520
            TabIndex        =   151
            Top             =   1860
            Width           =   5055
            Begin VB.TextBox txtCustom 
               Height          =   285
               Left            =   2400
               TabIndex        =   176
               Top             =   480
               Width           =   2415
            End
            Begin VB.OptionButton radFile_BDA 
               Caption         =   "Base - Date - Auto Number"
               Height          =   195
               Left            =   2640
               TabIndex        =   163
               Top             =   2520
               Width           =   2295
            End
            Begin VB.ComboBox cmbBase 
               Height          =   315
               Left            =   240
               TabIndex        =   162
               Top             =   480
               Width           =   2055
            End
            Begin VB.OptionButton radFile_BAD 
               Caption         =   "Base - Auto Number - Date"
               Height          =   255
               Left            =   240
               TabIndex        =   161
               Top             =   2520
               Width           =   2295
            End
            Begin VB.OptionButton radFile_DA 
               Caption         =   "Date - Auto Number"
               Height          =   195
               Left            =   2640
               TabIndex        =   160
               Top             =   2160
               Width           =   2175
            End
            Begin VB.OptionButton radFile_BD 
               Caption         =   "Base - Date"
               Height          =   255
               Left            =   2640
               TabIndex        =   159
               Top             =   1800
               Width           =   1455
            End
            Begin VB.OptionButton radFile_BA 
               Caption         =   "Base - Auto Number"
               Height          =   255
               Left            =   2640
               TabIndex        =   158
               Top             =   1440
               Width           =   2175
            End
            Begin VB.OptionButton radFile_D 
               Caption         =   "Date"
               Height          =   195
               Left            =   240
               TabIndex        =   157
               Top             =   2160
               Width           =   1095
            End
            Begin VB.OptionButton radFile_A 
               Caption         =   "Auto Number"
               Height          =   255
               Left            =   240
               TabIndex        =   156
               Top             =   1800
               Width           =   1695
            End
            Begin VB.OptionButton radFile_B 
               Caption         =   "Base"
               Height          =   255
               Left            =   240
               TabIndex        =   155
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label lblBase 
               Caption         =   "Custom Base"
               Height          =   255
               Left            =   2400
               TabIndex        =   177
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label23 
               Caption         =   "Base"
               Height          =   255
               Left            =   240
               TabIndex        =   154
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label22 
               Caption         =   "Current Format"
               Height          =   255
               Left            =   240
               TabIndex        =   153
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label lblFilePreview 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   240
               TabIndex        =   152
               Top             =   1080
               Width           =   4575
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Folder Name Generator Settings"
            Height          =   5415
            Left            =   -74760
            TabIndex        =   133
            Top             =   1500
            Width           =   3135
            Begin VB.OptionButton radFolder_BDF 
               Caption         =   "Base - Date - Folder Name"
               Height          =   315
               Left            =   120
               TabIndex        =   150
               Top             =   5040
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_BFD 
               Caption         =   "Base - Folder Name - Date"
               Height          =   255
               Left            =   120
               TabIndex        =   149
               Top             =   4800
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_DBF 
               Caption         =   "Date - Base -Folder Name"
               Height          =   375
               Left            =   120
               TabIndex        =   148
               Top             =   4440
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_FDB 
               Caption         =   "Folder Name - Date - Base"
               Height          =   255
               Left            =   120
               TabIndex        =   147
               Top             =   4200
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_DFB 
               Caption         =   "Date - Folder Name - Base"
               Height          =   375
               Left            =   120
               TabIndex        =   146
               Top             =   3840
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_FBD 
               Caption         =   "Folder Name - Base - Date"
               Height          =   255
               Left            =   120
               TabIndex        =   145
               Top             =   3600
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_FA 
               Caption         =   "Folder Name (Auto-Incrementing)"
               Height          =   375
               Left            =   120
               TabIndex        =   144
               Top             =   3240
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_BA 
               Caption         =   "Base (Auto-Incrementing)"
               Height          =   255
               Left            =   120
               TabIndex        =   143
               Top             =   3000
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_DF 
               Caption         =   "Date - Folder Name"
               Height          =   375
               Left            =   120
               TabIndex        =   142
               Top             =   2640
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_DB 
               Caption         =   "Date - Base"
               Height          =   255
               Left            =   120
               TabIndex        =   141
               Top             =   2400
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_BD 
               Caption         =   "Base - Date"
               Height          =   375
               Left            =   120
               TabIndex        =   140
               Top             =   2040
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_FD 
               Caption         =   "Folder Name - Date"
               Height          =   255
               Left            =   120
               TabIndex        =   139
               Top             =   1800
               Width           =   2895
            End
            Begin VB.OptionButton radFolder_D 
               Caption         =   "Date"
               Height          =   375
               Left            =   120
               TabIndex        =   138
               Top             =   1440
               Width           =   2895
            End
            Begin VB.TextBox txtFolderBase 
               Height          =   285
               Left            =   120
               TabIndex        =   136
               Top             =   450
               Width           =   2895
            End
            Begin VB.Label Label20 
               Caption         =   "Folder Name"
               Height          =   255
               Left            =   120
               TabIndex        =   137
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblFolderPreview 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   120
               TabIndex        =   135
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Label Label19 
               Caption         =   "Current Format"
               Height          =   255
               Left            =   120
               TabIndex        =   134
               Top             =   840
               Width           =   1215
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Seperator"
            Height          =   1095
            Left            =   -71520
            TabIndex        =   129
            Top             =   4860
            Width           =   1215
            Begin VB.OptionButton radSepChar_1 
               Caption         =   "_"
               Height          =   255
               Left            =   600
               TabIndex        =   131
               Top             =   720
               Width           =   495
            End
            Begin VB.OptionButton radSepChar_0 
               Caption         =   "-"
               Height          =   255
               Left            =   120
               TabIndex        =   130
               Top             =   720
               Width           =   375
            End
            Begin VB.Label Label18 
               Caption         =   "Seperating Character"
               Height          =   495
               Left            =   120
               TabIndex        =   132
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.CheckBox AutoSamplIDcheck 
            Caption         =   "Auto Increment Sample ID "
            Height          =   255
            Left            =   -70440
            TabIndex        =   128
            Top             =   2760
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   3480
            TabIndex        =   127
            Text            =   "Combo1"
            Top             =   6180
            Width           =   2895
         End
         Begin VB.TextBox lohmRegulatorText 
            Height          =   285
            Left            =   -74880
            TabIndex        =   124
            Text            =   "Text6"
            Top             =   4620
            Width           =   615
         End
         Begin VB.TextBox lohmFlowText 
            Height          =   285
            Left            =   -74880
            TabIndex        =   122
            Text            =   "Text6"
            Top             =   3660
            Width           =   615
         End
         Begin VB.Frame Frame14 
            Caption         =   "Permeability Logging"
            Height          =   1815
            Left            =   -74760
            TabIndex        =   117
            Top             =   5820
            Width           =   8175
            Begin VB.CommandButton permeabilityLoggingSelectButton 
               Caption         =   "Select File"
               Height          =   375
               Left            =   240
               TabIndex        =   121
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CheckBox permeabilityLoggingCheckBox 
               Caption         =   "Log permeability results"
               Height          =   255
               Left            =   240
               TabIndex        =   118
               Top             =   360
               Width           =   7695
            End
            Begin VB.Label permeabilityLoggingFileLabel 
               Height          =   255
               Left            =   360
               TabIndex        =   120
               Top             =   960
               Width           =   7695
            End
            Begin VB.Label Label16 
               Caption         =   "File:"
               Height          =   255
               Left            =   240
               TabIndex        =   119
               Top             =   720
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Test Endpoints"
            Height          =   1455
            Left            =   -74760
            TabIndex        =   96
            Top             =   1140
            Width           =   5895
            Begin VB.OptionButton endpointoption 
               Caption         =   "Based on pressure"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   12
               Top             =   360
               Value           =   -1  'True
               Width           =   4935
            End
            Begin VB.OptionButton endpointoption 
               Caption         =   "Based on pore diameter (capflow and bubble point tests only)"
               Height          =   615
               Index           =   2
               Left            =   240
               TabIndex        =   13
               Top             =   720
               Width           =   5535
            End
         End
         Begin VB.CheckBox autoincCheck 
            Caption         =   "Auto increment data file names"
            Height          =   255
            Left            =   -74520
            TabIndex        =   15
            Top             =   3120
            Width           =   3735
         End
         Begin VB.CheckBox hidePromptsCheck 
            Caption         =   "Hide sample load prompts"
            Height          =   255
            Left            =   -74520
            TabIndex        =   16
            Top             =   3480
            Width           =   3615
         End
         Begin VB.CheckBox minPressCheck 
            Caption         =   "Use min. pressure in dry curve"
            Height          =   255
            Left            =   -74520
            TabIndex        =   17
            Top             =   3840
            Width           =   3615
         End
         Begin VB.CheckBox advancedCheck 
            Caption         =   "Use advanced settings only"
            Height          =   255
            Left            =   -74520
            TabIndex        =   14
            Top             =   2760
            Width           =   5655
         End
         Begin VB.Frame Frame4 
            Caption         =   "Units"
            Height          =   5175
            Left            =   120
            TabIndex        =   90
            Top             =   1140
            Width           =   2415
            Begin VB.ComboBox pressCombo 
               Height          =   315
               ItemData        =   "Prefs.frx":01A8
               Left            =   120
               List            =   "Prefs.frx":01AA
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   720
               Width           =   1215
            End
            Begin VB.ComboBox lengthCombo 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   1680
               Width           =   1215
            End
            Begin VB.ComboBox thickCombo 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   2640
               Width           =   1215
            End
            Begin VB.ComboBox densityCombo 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   3600
               Width           =   1215
            End
            Begin VB.ComboBox massCombo 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   4560
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Pressure"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label Label2 
               Caption         =   "Length"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   1320
               Width           =   1935
            End
            Begin VB.Label Label3 
               Caption         =   "Thickness"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   2280
               Width           =   1935
            End
            Begin VB.Label label4 
               Caption         =   "Density"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   3240
               Width           =   1935
            End
            Begin VB.Label label5 
               Caption         =   "Mass"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   4200
               Width           =   1935
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Comm Port"
            Height          =   975
            Left            =   2640
            TabIndex        =   89
            Top             =   1140
            Width           =   2295
            Begin VB.ComboBox commPortCombo 
               Height          =   315
               Left            =   255
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   375
               Width           =   1815
            End
         End
         Begin VB.CommandButton setFontButton 
            Caption         =   "Set Font Properties"
            Height          =   375
            Left            =   3720
            TabIndex        =   11
            Top             =   4740
            Width           =   2415
         End
         Begin VB.Frame Frame1 
            Caption         =   "Log File"
            Height          =   2175
            Left            =   2640
            TabIndex        =   86
            Top             =   2340
            Width           =   4575
            Begin VB.CommandButton changeFileButton 
               Caption         =   "Change File"
               Height          =   375
               Left            =   1080
               TabIndex        =   10
               Top             =   1680
               Width           =   2415
            End
            Begin VB.CheckBox logCheck 
               Caption         =   "Enable logging"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   360
               Width           =   4335
            End
            Begin VB.Label logText 
               Alignment       =   2  'Center
               Caption         =   "c:\Program Files\CapWin\log.txt"
               Height          =   615
               Left            =   120
               TabIndex        =   88
               Top             =   960
               Width           =   4335
            End
            Begin VB.Label Label6 
               Caption         =   "Log file:"
               Height          =   255
               Left            =   480
               TabIndex        =   87
               Top             =   720
               Width           =   1215
            End
         End
         Begin VB.TextBox pholdFreq 
            Height          =   315
            Left            =   -71640
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox numAvePoints 
            Height          =   315
            Left            =   -71640
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   1200
            Width           =   735
         End
         Begin VB.Frame Frame3 
            Caption         =   "Pass/fail options"
            Height          =   2835
            Left            =   -74760
            TabIndex        =   80
            Top             =   3000
            Width           =   6855
            Begin VB.TextBox holdRate 
               Height          =   285
               Left            =   1440
               TabIndex        =   29
               Text            =   "Text1"
               Top             =   360
               Width           =   855
            End
            Begin VB.OptionButton pressHoldUnits 
               Caption         =   "PSI/sec"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   30
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton pressHoldUnits 
               Caption         =   "PSI/min"
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   31
               Top             =   600
               Width           =   1455
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Abort if test fails"
               Height          =   255
               Left            =   360
               TabIndex        =   32
               Top             =   960
               Width           =   5895
            End
            Begin VB.Frame Frame6 
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   375
               Left            =   240
               TabIndex        =   82
               Top             =   1560
               Width           =   6495
               Begin VB.OptionButton failDPDT 
                  Caption         =   "Delta P / Delta t"
                  Height          =   375
                  Index           =   0
                  Left            =   480
                  TabIndex        =   33
                  Top             =   0
                  Width           =   2775
               End
               Begin VB.OptionButton failDPDT 
                  Caption         =   "Delta P only"
                  Height          =   375
                  Index           =   1
                  Left            =   3360
                  TabIndex        =   34
                  Top             =   0
                  Width           =   2895
               End
            End
            Begin VB.Frame Frame10 
               BorderStyle     =   0  'None
               Caption         =   "Frame10"
               Height          =   495
               Left            =   240
               TabIndex        =   81
               Top             =   2160
               Width           =   6375
               Begin VB.OptionButton regress 
                  Caption         =   "Linear regression"
                  Height          =   255
                  Index           =   1
                  Left            =   3360
                  TabIndex        =   73
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.OptionButton regress 
                  Caption         =   "Two-point determination"
                  Height          =   255
                  Index           =   0
                  Left            =   480
                  TabIndex        =   67
                  Top             =   120
                  Width           =   2775
               End
            End
            Begin VB.Label pLabel 
               Alignment       =   1  'Right Justify
               Caption         =   "Fail rate:"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   85
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label pLabel 
               Caption         =   "Base failure on"
               Height          =   255
               Index           =   5
               Left            =   360
               TabIndex        =   84
               Top             =   1320
               Width           =   5775
            End
            Begin VB.Label pLabel 
               Caption         =   "Using:"
               Height          =   255
               Index           =   7
               Left            =   360
               TabIndex        =   83
               Top             =   1920
               Width           =   2175
            End
         End
         Begin VB.TextBox delayTime 
            Height          =   315
            Left            =   -74520
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox testLength 
            Height          =   315
            Left            =   -74520
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   1200
            Width           =   735
         End
         Begin VB.Frame Frame7 
            Caption         =   "Graph options"
            Height          =   1455
            Left            =   -74760
            TabIndex        =   76
            Top             =   5820
            Width           =   6855
            Begin VB.CheckBox autoscaleCheck 
               Caption         =   "Autoscale Y-axis"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   240
               Width           =   4215
            End
            Begin VB.Frame Frame8 
               BorderStyle     =   0  'None
               Caption         =   "Frame8"
               Height          =   855
               Left            =   240
               TabIndex        =   77
               Top             =   480
               Width           =   6495
               Begin VB.TextBox minY 
                  Height          =   285
                  Left            =   0
                  TabIndex        =   37
                  Text            =   "Text1"
                  Top             =   120
                  Width           =   855
               End
               Begin VB.TextBox maxY 
                  Height          =   285
                  Left            =   0
                  TabIndex        =   38
                  Text            =   "Text2"
                  Top             =   480
                  Width           =   855
               End
               Begin VB.Label Label8 
                  Caption         =   "Minimum Y value"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   79
                  Top             =   120
                  Width           =   4455
               End
               Begin VB.Label Label9 
                  Caption         =   "Maximum Y value"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   78
                  Top             =   480
                  Width           =   4455
               End
            End
         End
         Begin VB.TextBox lohmPercent 
            Height          =   285
            Left            =   -74880
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   1980
            Width           =   615
         End
         Begin VB.TextBox lpMintime 
            Height          =   285
            Left            =   -74880
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   1260
            Width           =   735
         End
         Begin VB.Frame Frame9 
            Caption         =   "Show results at end of test"
            Height          =   1935
            Left            =   -74760
            TabIndex        =   75
            Top             =   5460
            Width           =   5895
            Begin VB.OptionButton resultOption 
               Caption         =   "No results"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   21
               Top             =   360
               Width           =   5535
            End
            Begin VB.OptionButton resultOption 
               Caption         =   "Summary sheet in CapRep"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   22
               Top             =   720
               Width           =   5535
            End
            Begin VB.OptionButton resultOption 
               Caption         =   "Run report in CapRep"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   23
               Top             =   1080
               Width           =   5535
            End
            Begin VB.OptionButton resultOption 
               Caption         =   "Calculate results in CapWin"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   24
               Top             =   1440
               Value           =   -1  'True
               Width           =   5535
            End
         End
         Begin VB.CheckBox microflowregulatorcheck 
            Caption         =   "Control regulator to compensate for back pressure"
            Height          =   255
            Left            =   -74640
            TabIndex        =   43
            Top             =   1320
            Width           =   7935
         End
         Begin VB.CheckBox norefillcheck 
            Caption         =   "Do not refill penetrometer in elevated pressure test"
            Height          =   495
            Left            =   -74880
            TabIndex        =   40
            Top             =   1680
            Width           =   8175
         End
         Begin VB.CheckBox delaycompressionliquidcheck 
            Caption         =   "Delay compression of sample until after initial fill of liquid"
            Height          =   495
            Left            =   -74880
            TabIndex        =   41
            Top             =   2280
            Width           =   8175
         End
         Begin VB.CheckBox linSealCheck 
            Caption         =   "Use linear seal length for flow calculations"
            Height          =   255
            Left            =   -74640
            TabIndex        =   44
            Top             =   1800
            Width           =   7695
         End
         Begin VB.CheckBox mf_settle_check 
            Caption         =   "Delay test until stable pressure in microflow volume"
            Height          =   255
            Left            =   -74640
            TabIndex        =   45
            Top             =   2280
            Width           =   8175
         End
         Begin VB.TextBox mf_settle_pressure_text 
            Height          =   285
            Left            =   -74400
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   2640
            Width           =   855
         End
         Begin VB.TextBox mf_settle_time_text 
            Height          =   285
            Left            =   -74400
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   3000
            Width           =   855
         End
         Begin VB.CheckBox mf_temperature_check 
            Caption         =   "Save temperature data to test data file"
            Height          =   255
            Left            =   -74640
            TabIndex        =   48
            Top             =   3480
            Width           =   7815
         End
         Begin VB.CheckBox curveFitCheck 
            Caption         =   "Automatically curve-fit data file at end of test"
            Height          =   255
            Left            =   -74520
            TabIndex        =   20
            Top             =   4980
            Width           =   7455
         End
         Begin VB.TextBox cfNumPoints 
            Height          =   375
            Left            =   -74760
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   1380
            Width           =   855
         End
         Begin VB.TextBox cfPercentError 
            Height          =   375
            Left            =   -74760
            TabIndex        =   50
            Text            =   "Text2"
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox cfMaxPSI 
            Height          =   375
            Left            =   -74760
            TabIndex        =   51
            Text            =   "Text3"
            Top             =   2580
            Width           =   855
         End
         Begin VB.TextBox minFlowText 
            Height          =   285
            Left            =   -73920
            TabIndex        =   19
            Text            =   "0"
            Top             =   4620
            Width           =   1455
         End
         Begin VB.CheckBox minFlowCheck 
            Caption         =   "Use min. flow in dry curve"
            Height          =   255
            Left            =   -74520
            TabIndex        =   18
            Top             =   4260
            Width           =   3495
         End
         Begin VB.TextBox lohmTimeoutText 
            Height          =   285
            Left            =   -74880
            TabIndex        =   74
            Text            =   "Text6"
            Top             =   2700
            Width           =   615
         End
         Begin VB.CheckBox zeroTempCheck 
            Caption         =   "Turn off heater when test is done (non-circulating systems only)"
            Height          =   255
            Left            =   -74280
            TabIndex        =   59
            Top             =   1620
            Width           =   6975
         End
         Begin VB.Frame Frame11 
            Caption         =   "Single-point and Multi-Set Tests"
            Height          =   2895
            Left            =   -74760
            TabIndex        =   66
            Top             =   1020
            Width           =   8175
            Begin VB.TextBox txtNumberdp 
               Height          =   375
               Left            =   120
               TabIndex        =   187
               Text            =   "10"
               Top             =   2040
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox txtSetDura 
               Height          =   375
               Left            =   120
               TabIndex        =   186
               Text            =   "60"
               Top             =   1560
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox txtNumSets 
               Height          =   375
               Left            =   120
               TabIndex        =   185
               Text            =   "3"
               Top             =   1080
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox txtTargVol 
               Height          =   375
               Left            =   120
               TabIndex        =   184
               Text            =   "1"
               Top             =   600
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox Text5 
               Height          =   375
               Left            =   1800
               TabIndex        =   56
               Text            =   "1"
               Top             =   2400
               Width           =   735
            End
            Begin VB.TextBox Text4 
               Height          =   375
               Left            =   1800
               TabIndex        =   55
               Text            =   "1"
               Top             =   1920
               Width           =   735
            End
            Begin VB.TextBox Text3 
               Height          =   375
               Left            =   1800
               TabIndex        =   54
               Text            =   "1"
               Top             =   1440
               Width           =   735
            End
            Begin VB.TextBox Text2 
               Height          =   375
               Left            =   1800
               TabIndex        =   53
               Text            =   "10"
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox Text1 
               Height          =   375
               Left            =   1800
               TabIndex        =   52
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label24 
               Caption         =   "Multi Set Options"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   188
               Top             =   240
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.Label NumberofDatapoints 
               Caption         =   "Number of data points per set"
               Height          =   255
               Left            =   720
               TabIndex        =   183
               Top             =   2160
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.Label SetDuration 
               Caption         =   "Duration of each Set in Seconds"
               Height          =   255
               Left            =   720
               TabIndex        =   182
               Top             =   1680
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.Label SetTotal 
               Caption         =   "Total Number of Sets "
               Height          =   255
               Left            =   720
               TabIndex        =   181
               Top             =   1200
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.Label targetVol 
               Caption         =   "Target Volocity l/s"
               Height          =   255
               Left            =   720
               TabIndex        =   180
               Top             =   720
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.Label Label21 
               Caption         =   "Single-point Test Options"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2760
               TabIndex        =   179
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label Label11 
               Caption         =   "Target pressure"
               Height          =   255
               Index           =   5
               Left            =   2880
               TabIndex        =   72
               Top             =   600
               Width           =   2655
            End
            Begin VB.Label Label11 
               Caption         =   "Number of readings taken for each point"
               Height          =   255
               Index           =   3
               Left            =   2880
               TabIndex        =   71
               Top             =   2520
               Width           =   5055
            End
            Begin VB.Label Label11 
               Caption         =   "Delay time at start of test (seconds)"
               Height          =   255
               Index           =   2
               Left            =   2880
               TabIndex        =   70
               Top             =   1080
               Width           =   4695
            End
            Begin VB.Label Label11 
               Caption         =   "Time between data points (seconds)"
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   69
               Top             =   2040
               Width           =   4695
            End
            Begin VB.Label Label11 
               Caption         =   "Test duration (seconds)"
               Height          =   255
               Index           =   0
               Left            =   2880
               TabIndex        =   68
               Top             =   1560
               Width           =   2535
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Averaging test"
            Height          =   1215
            Left            =   -74760
            TabIndex        =   63
            Top             =   4020
            Width           =   8175
            Begin VB.CheckBox GP_avgTestCheck 
               Caption         =   "Run gas permeability in averaging mode"
               Height          =   255
               Left            =   240
               TabIndex        =   57
               Top             =   360
               Width           =   7695
            End
            Begin VB.TextBox numAvgTests 
               Height          =   285
               Left            =   240
               TabIndex        =   58
               Text            =   "Text6"
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label15 
               Caption         =   "Number of tests to average"
               Height          =   255
               Left            =   960
               TabIndex        =   65
               Top             =   720
               Width           =   5655
            End
         End
         Begin VB.CheckBox BPTLCheck 
            Caption         =   "Bubble Point Time Log"
            Height          =   255
            Left            =   -74640
            TabIndex        =   60
            Top             =   2460
            Width           =   7695
         End
         Begin VB.TextBox BPTLSecondsText 
            Height          =   285
            Left            =   -74400
            TabIndex        =   61
            Top             =   2820
            Width           =   735
         End
         Begin VB.TextBox BPTLMaxText 
            Height          =   285
            Left            =   -74400
            TabIndex        =   62
            Top             =   3300
            Width           =   735
         End
         Begin VB.CheckBox BubblerCheck 
            Caption         =   "Use Bubbler"
            Height          =   255
            Left            =   -74640
            TabIndex        =   64
            Top             =   4020
            Width           =   4455
         End
         Begin VB.Label minEndPsiText 
            Caption         =   "Minimum ending PSI"
            Height          =   255
            Left            =   -73680
            TabIndex        =   274
            Top             =   2640
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label and2 
            Caption         =   "Minimum ending PSI"
            Height          =   255
            Left            =   -70800
            TabIndex        =   273
            Top             =   2640
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label and1 
            Caption         =   "Maximum starting PSI"
            Height          =   255
            Left            =   -70800
            TabIndex        =   269
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label startingPText 
            Caption         =   "Minimum starting PSI"
            Height          =   255
            Left            =   -73680
            TabIndex        =   267
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblV6LPRegIncWait 
            Caption         =   "Regulator increase wait period in milliseconds"
            Height          =   255
            Left            =   -74040
            TabIndex        =   266
            Top             =   6000
            Width           =   7095
         End
         Begin VB.Label lblV6LPRegIncAmount 
            Caption         =   "Regulator increase amount"
            Height          =   255
            Left            =   -74040
            TabIndex        =   264
            Top             =   5640
            Width           =   7095
         End
         Begin VB.Label pLabel 
            Caption         =   "Drain time (seconds)"
            Height          =   255
            Index           =   9
            Left            =   -73680
            TabIndex        =   262
            Top             =   4680
            Width           =   6495
         End
         Begin VB.Label lblLPFlushPressure 
            Caption         =   "Pressure for flushing"
            Height          =   255
            Left            =   -73680
            TabIndex        =   257
            Top             =   3720
            Width           =   7095
         End
         Begin VB.Label lblLPFlushCCs 
            Caption         =   "Number of cc's to flush"
            Height          =   255
            Left            =   -73680
            TabIndex        =   255
            Top             =   3360
            Width           =   7095
         End
         Begin VB.Label Label49 
            Caption         =   "Account Type:"
            Height          =   255
            Left            =   -71760
            TabIndex        =   249
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblUsers 
            Caption         =   "Users"
            Height          =   255
            Left            =   -74400
            TabIndex        =   246
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lblPumpSpeedDesc 
            Caption         =   "(0-255) determines how fast the pump will pump liquid."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -73320
            TabIndex        =   243
            Top             =   4800
            Width           =   5055
         End
         Begin VB.Label lblPumpSpeed 
            Caption         =   "Pump Speed"
            Height          =   255
            Left            =   -74160
            TabIndex        =   241
            Top             =   4560
            Width           =   975
         End
         Begin VB.Label time4Label 
            Caption         =   "seconds to reverse the flow of fluid from the wetting tube."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -72720
            TabIndex        =   240
            Top             =   2280
            Width           =   5295
         End
         Begin VB.Label pLabel 
            Caption         =   "The number of consecutive points used to detect a bubble point"
            Height          =   255
            Index           =   8
            Left            =   -74040
            TabIndex        =   238
            Top             =   1080
            Width           =   7095
         End
         Begin VB.Label lohmCaption5 
            Caption         =   "Lohm Tolerance (only values below this will be recorded to the lohm table)"
            Height          =   375
            Left            =   -73920
            TabIndex        =   236
            Top             =   5400
            Width           =   6975
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            Caption         =   "Minimum Adjustment Flow (cc)  "
            Height          =   255
            Left            =   -74160
            TabIndex        =   232
            Top             =   6720
            Width           =   2295
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   "Initial Humidity Wait Time (sec)  "
            Height          =   255
            Left            =   -74160
            TabIndex        =   230
            Top             =   6240
            Width           =   2295
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "Stability Sleep Time (msec)  "
            Height          =   255
            Left            =   -74160
            TabIndex        =   229
            Top             =   5760
            Width           =   2295
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "Stability Tolerance (%)  "
            Height          =   255
            Left            =   -74280
            TabIndex        =   228
            Top             =   5280
            Width           =   2415
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Min. Stability Wait Time (sec)  "
            Height          =   255
            Left            =   -74160
            TabIndex        =   227
            Top             =   4800
            Width           =   2295
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Max. Stability Wait Time (sec)  "
            Height          =   255
            Left            =   -74400
            TabIndex        =   226
            Top             =   4320
            Width           =   2535
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Target Tolerance (%)  "
            Height          =   255
            Left            =   -74280
            TabIndex        =   225
            Top             =   3840
            Width           =   2415
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Min. Adjustment Wait Time (sec)  "
            Height          =   255
            Left            =   -74640
            TabIndex        =   224
            Top             =   3360
            Width           =   2775
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Max. Adjustment Wait Time (sec)  "
            Height          =   255
            Left            =   -74520
            TabIndex        =   219
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label Label26 
            Caption         =   "volume of fluid to flow onto the sample."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -72720
            TabIndex        =   214
            Top             =   2760
            Width           =   5295
         End
         Begin VB.Label lblTargetHumidity 
            Alignment       =   1  'Right Justify
            Caption         =   "Target Humidity (%)  "
            Height          =   255
            Left            =   -74160
            TabIndex        =   210
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label lblInfoHeaders 
            Caption         =   "Information Headers"
            Height          =   255
            Left            =   -73920
            TabIndex        =   207
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label lblNumOfInfoLines 
            Caption         =   "Number of information lines"
            Height          =   255
            Left            =   -74400
            TabIndex        =   204
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label25 
            Caption         =   "Temperature at which the door lock is released, if equipped"
            Height          =   255
            Left            =   -73440
            TabIndex        =   201
            Top             =   1920
            Width           =   5655
         End
         Begin VB.Label test_unsupport 
            Caption         =   "The Automatic Wet Feature is not supported by your hardware configuration or the hardware is not turned ON."
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1335
            Left            =   -74400
            TabIndex        =   199
            Top             =   5880
            Visible         =   0   'False
            Width           =   7215
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label29 
            Caption         =   "seconds to have fluid drain out of the test chamber."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -73320
            TabIndex        =   198
            Top             =   4080
            Width           =   5055
         End
         Begin VB.Label drain_time 
            Caption         =   "Drain Time"
            Height          =   255
            Left            =   -74160
            TabIndex        =   196
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label28 
            Caption         =   "seconds to have the sample soak in fluid."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -73320
            TabIndex        =   195
            Top             =   3480
            Width           =   4575
         End
         Begin VB.Label Label27 
            Caption         =   "seconds to enable the flow of fluid onto the test sample."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -72720
            TabIndex        =   194
            Top             =   1800
            Width           =   5295
         End
         Begin VB.Label soak_time 
            Caption         =   "Soak Time"
            Height          =   255
            Left            =   -74160
            TabIndex        =   193
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label wet_time 
            Caption         =   "Wet Time"
            Height          =   255
            Left            =   -74160
            TabIndex        =   192
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "Languages"
            Height          =   375
            Left            =   3600
            TabIndex        =   126
            Top             =   5460
            Width           =   2655
         End
         Begin VB.Label lohmCaption4 
            Caption         =   "Lohm regulator increase factor.  1=normal.  2=twice as fast"
            Height          =   615
            Left            =   -73920
            TabIndex        =   125
            Top             =   4620
            Width           =   6975
         End
         Begin VB.Label lohmCaption3 
            Caption         =   "Allowable flow change in 5 seconds during lohm stability (in cc/min)"
            Height          =   615
            Left            =   -73920
            TabIndex        =   123
            Top             =   3660
            Width           =   6975
         End
         Begin VB.Label pLabel 
            Caption         =   "Length of test (sec)"
            Height          =   255
            Index           =   0
            Left            =   -73680
            TabIndex        =   116
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label pLabel 
            Caption         =   "Delay time (sec)"
            Height          =   255
            Index           =   1
            Left            =   -73680
            TabIndex        =   115
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label pLabel 
            Caption         =   "Number of points used for data averaging"
            Height          =   255
            Index           =   2
            Left            =   -70800
            TabIndex        =   114
            Top             =   1200
            Width           =   3855
         End
         Begin VB.Label pLabel 
            Caption         =   "Seconds to wait between readings"
            Height          =   255
            Index           =   3
            Left            =   -70800
            TabIndex        =   113
            Top             =   1680
            Width           =   3735
         End
         Begin VB.Label Label12 
            Caption         =   "Lohm Calibration"
            Height          =   255
            Left            =   -74880
            TabIndex        =   112
            Top             =   1620
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   "%"
            Height          =   255
            Left            =   -74280
            TabIndex        =   111
            Top             =   1980
            Width           =   135
         End
         Begin VB.Label pLabel 
            Caption         =   "Minimum time (seconds) between readings"
            Height          =   255
            Index           =   6
            Left            =   -74040
            TabIndex        =   110
            Top             =   1260
            Width           =   7095
         End
         Begin VB.Label mf_settle_label1 
            Caption         =   "Maximum Allowable Pressure Change - PSI"
            Height          =   255
            Left            =   -73440
            TabIndex        =   109
            Top             =   2640
            Width           =   6855
         End
         Begin VB.Label mf_settle_label2 
            Caption         =   "Time during which pressure must be stable - Seconds"
            Height          =   255
            Left            =   -73440
            TabIndex        =   108
            Top             =   3000
            Width           =   6855
         End
         Begin VB.Label Label10 
            Caption         =   "placeholder - filled in form_load"
            Height          =   2775
            Left            =   -74760
            TabIndex        =   107
            Top             =   3780
            Width           =   8295
         End
         Begin VB.Label cflabel 
            Caption         =   "Number of points used for fitting (3 - 10)"
            Height          =   375
            Index           =   0
            Left            =   -73680
            TabIndex        =   106
            Top             =   1380
            Width           =   5775
         End
         Begin VB.Label cflabel 
            Caption         =   "Maximum allowable percentage error"
            Height          =   375
            Index           =   1
            Left            =   -73680
            TabIndex        =   105
            Top             =   1980
            Width           =   5775
         End
         Begin VB.Label cflabel 
            Caption         =   "Maximum pressure difference between two points"
            Height          =   375
            Index           =   2
            Left            =   -73680
            TabIndex        =   104
            Top             =   2580
            Width           =   5775
         End
         Begin VB.Label Label13 
            Caption         =   "l/min"
            Height          =   255
            Left            =   -72240
            TabIndex        =   103
            Top             =   4620
            Width           =   2895
         End
         Begin VB.Label lohmCaption 
            Caption         =   "Percentage of maximum flow at which lohm calibration begins"
            Height          =   615
            Left            =   -73920
            TabIndex        =   102
            Top             =   1980
            Width           =   6975
         End
         Begin VB.Label lohmCaption2 
            Caption         =   "Lohm timeout value (number of cycles before calibration stops if the flow cannot be increased)"
            Height          =   615
            Left            =   -73920
            TabIndex        =   101
            Top             =   2700
            Width           =   6975
         End
         Begin VB.Label Label14 
            Caption         =   "Temperature control"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74640
            TabIndex        =   100
            Top             =   1260
            Width           =   7095
         End
         Begin VB.Line Line1 
            X1              =   -74640
            X2              =   -67800
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label Label11 
            Caption         =   $"Prefs.frx":01AC
            Height          =   495
            Index           =   4
            Left            =   -74640
            TabIndex        =   99
            Top             =   5340
            Width           =   7815
         End
         Begin VB.Label BPTLSecondsLabel 
            Caption         =   "Seconds between readings"
            Height          =   255
            Left            =   -73440
            TabIndex        =   98
            Top             =   2820
            Width           =   6375
         End
         Begin VB.Label BPTLMaxLabel 
            Caption         =   "Maximum number of data points"
            Height          =   255
            Left            =   -73440
            TabIndex        =   97
            Top             =   3300
            Width           =   6375
         End
         Begin VB.Line Line2 
            X1              =   -74640
            X2              =   -67800
            Y1              =   3780
            Y2              =   3780
         End
      End
   End
End
Attribute VB_Name = "prefsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type DCB
    DCBlength As Long
    BaudRate As Long
    fBitFields As Long 'See Comments in Win32API.Txt
    wReserved As Integer
    XonLim As Integer
    XoffLim As Integer
    ByteSize As Byte
    Parity As Byte
    StopBits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte
    wReserved1 As Integer 'Reserved; Do Not Use
End Type

Private Type COMMCONFIG
    dwSize As Long
    wVersion As Integer
    wReserved As Integer
    dcbx As DCB
    dwProviderSubType As Long
    dwProviderOffset As Long
    dwProviderSize As Long
    wcProviderData As Byte
End Type

'these are used for making sure the starting and ending pressures are within
'certain parameters during the test.
Dim startP1 As Double
Dim startP2 As Double
Dim endP1 As Double
Dim endP2 As Double
Dim startend As Integer
Dim testString As String

Private Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long
Dim languageIndex As Integer
Dim originalLanguageIndex As Integer
Dim ts$(19)                     ' Text strings for this form
Dim UAC_Changed As Boolean

'added 11-21-07 --Denis
Private Sub auto_wet_check_Click()
    'Place code here when clicking to enable/diable the auto wet sample test.
    '"Auto wet test" tab options
    
    If auto_wet_check.value = 0 Then
        ' we don't write out changes until they click on the "save" button
        'WPPS "Capstuff", "auto_wet_used", "0", CSFile$
        
     '1st option
        optWetTime.Visible = False
        wet_time.Visible = False
        time1.Visible = False
        Label27.Visible = False
        time4.Visible = False
        time4Label.Visible = False
        txtPumpSpeed.Visible = False
        lblPumpSpeed.Visible = False
        lblPumpSpeedDesc.Visible = False
     '2nd option
        optWetVolume.Visible = False
        txtAutoWetVolume.Visible = False
        Label26.Visible = False
        soak_time.Visible = False
        time2.Visible = False
        Label28.Visible = False
     '3rd option
        drain_time.Visible = False
        time3.Visible = False
        Label29.Visible = False
    Else
        'WPPS "Capstuff", "auto_wet_used", "1", CSFile$
     '1st option
        optWetTime.Visible = True
        wet_time.Visible = True
        time1.Visible = True
        time4.Visible = True
        time4Label.Visible = True
        optWetVolume.Visible = True
        Label26.Visible = True
        txtAutoWetVolume.Visible = True
        ' the text boxes will already have been filled in in the form load
        'time1.Text = gpps2(Curr_U$, "wet_time", IFile$, "0")
        Label27.Visible = True
     '2nd option
        soak_time.Visible = True
        time2.Visible = True
        'time2.Text = gpps2(Curr_U$, "soak_time", IFile$, "0")
        Label28.Visible = True
     '3rd option
        drain_time.Visible = True
        time3.Visible = True
        'time3.Text = gpps2(Curr_U$, "drain_time", IFile$, "0")
        Label29.Visible = True
        txtPumpSpeed.Visible = True
        lblPumpSpeed.Visible = True
        lblPumpSpeedDesc.Visible = True
    End If
End Sub

Private Sub changeFileButton_Click()

    fsel_name$ = logText.Caption
    fsel_title$ = ts$(9)                ' "Set Logging File"
    fsel_io = False ' file doesn't have to exist
    fsel_path = "*.txt"
    fsel Me.hwnd
    If fsel_return$ <> "" Then
        logText.Caption = fsel_return$
    End If
    
End Sub

Private Sub autoscaleCheck_Click()

    Frame8.Visible = (autoscaleCheck.value = vbUnchecked)

End Sub

Private Sub chkEnableAdditionalInfo_Click()
    If chkEnableAdditionalInfo.value = 1 Then
        lblNumOfInfoLines.Visible = True
        txtNumberOfInfoLines.Visible = True
        lblInfoHeaders.Visible = True
        lstInfoLines.Visible = True
    Else
        lblNumOfInfoLines.Visible = False
        txtNumberOfInfoLines.Visible = False
        lblInfoHeaders.Visible = False
        lstInfoLines.Visible = False
    End If
End Sub

Private Sub chkEnableUAC_Click()
    Dim newPwd1 As String
    Dim newPwd2 As String
    
    If chkEnableUAC.value = vbChecked Then
        If lstUsers.ListCount = 0 Then
            MsgBox "This is the first time you are enabling the User Access Control system." + vbCrLf + _
                   "The software is automatically generating an Admin account.  Please " + vbCrLf + _
                   "enter a password for this account.", vbInformation
            newPwd1 = InputBox("Enter Admin Password", "New Password", "")
            newPwd2 = InputBox("Re-enter Admin Password", "New Password", "")
            While newPwd1 <> newPwd2
                MsgBox "Password mismatch.  Please re-enter.", vbInformation
                newPwd1 = InputBox("Enter Admin Password", "New Password", "")
                newPwd2 = InputBox("Re-enter Admin Password", "New Password", "")
            Wend
            lstUsers.AddItem "Admin"
            UAC_UserCount = 1
            ReDim UAC_Users(UAC_UserCount)
            UAC_Users(1).Username = "Admin"
            UAC_Users(1).Password = newPwd1
            UAC_Users(1).accessLevel = 0
        End If
    End If
    UAC_Enabled = chkEnableUAC.value
    UAC_Changed = True
End Sub

Private Sub cmdAddUser_Click()
    Dim newUser As String
    Dim newPwd1 As String
    Dim newPwd2 As String
    Dim i As Integer
    Dim userExists As Boolean
    
    If UAC_Enabled Then
        newUser = InputBox("Please enter a username for the new account", "Add User", "")
        If Trim(newUser) <> "" Then
            userExists = False
            For i = 1 To UAC_UserCount
                If newUser = UAC_Users(i).Username Then userExists = True
            Next i
            
            If userExists Then
                MsgBox "The user account " + newUser + " already exists.", vbInformation
            Else
                newPwd1 = InputBox("Please enter a password for the new account", "Add User", "")
                newPwd2 = InputBox("Please re-enter the password for the new account", "Add User", "")
                While newPwd1 <> newPwd2
                    MsgBox "Password mismatch.  Please re-enter.", vbInformation
                    newPwd1 = InputBox("Please enter a password for the new account", "Add User", "")
                    newPwd2 = InputBox("Please re-enter the password for the new account", "Add User", "")
                Wend
                UAC_UserCount = UAC_UserCount + 1
                ReDim Preserve UAC_Users(UAC_UserCount)
                UAC_Users(UAC_UserCount).Username = newUser
                UAC_Users(UAC_UserCount).Password = newPwd1
                UAC_Users(UAC_UserCount).accessLevel = 0
                lstUsers.AddItem newUser
                lstUsers.ListIndex = lstUsers.ListCount - 1
                UAC_Changed = True
            End If
        End If
    Else
        MsgBox "Please enable User Access Control first.", vbInformation
    End If
End Sub

Private Sub cmdChangePassword_Click()
    Dim lstIndex As Integer
    Dim newPwd1 As String
    Dim newPwd2 As String
    
    If UAC_Enabled Then
        lstIndex = lstUsers.ListIndex + 1
        If lstIndex > 0 Then
            newPwd1 = InputBox("Enter a new password", "New Password", "")
            newPwd2 = InputBox("Re-enter the password", "New Password", "")
            If newPwd1 = newPwd2 Then
                UAC_Users(lstIndex).Password = newPwd1
                MsgBox "Password changed.", vbInformation
                UAC_Changed = True
            Else
                MsgBox "Passwords were not the same.  Aborting password change.", vbInformation
            End If
        Else
            MsgBox "You must select an account first.", vbInformation
        End If
    End If
End Sub

Private Sub cmdRemoveUser_Click()
    Dim msgres As VbMsgBoxResult
    Dim lstIndex As Integer
    Dim i As Integer
    
    If UAC_Enabled Then
        lstIndex = lstUsers.ListIndex + 1
        If lstIndex > 1 Then
            If lstIndex = UAC_CurrentUser Then
                MsgBox "You cannot remove the currently logged in account", vbInformation
            Else
                msgres = MsgBox("Are you sure you want to remove the following account: " + UAC_Users(lstUsers.ListIndex + 1).Username, vbYesNo)
                If msgres = vbYes Then
                    For i = lstIndex + 1 To UAC_UserCount
                        UAC_Users(i - 1) = UAC_Users(i)
                    Next i
                    UAC_UserCount = UAC_UserCount - 1
                    ReDim Preserve UAC_Users(UAC_UserCount)
                    lstUsers.RemoveItem lstIndex - 1
                    UAC_Changed = True
                    MsgBox "Account removed.", vbInformation
                End If
            End If
        ElseIf lstIndex = 1 Then
            MsgBox "You cannot remove the Admin account.", vbInformation
        Else
            MsgBox "You must select an account to remove first.", vbInformation
        End If
    End If
End Sub

Private Sub Combo1_Click()
 Dim i As Integer
 
    If Combo1.Tag = "system" Then Exit Sub   ' Don't trigger when form is loading
    
    If Combo1.ListIndex <> languageIndex Then
        languageIndex = Combo1.ListIndex
        For i = 1 To languageCount
            If available_languages$(i, 2) = Combo1.Text Then
                language$ = EXE_Path + "languages\" + available_languages$(i, 1)
            End If
        Next i

        LoadTextStrings
     
        CAPFLOW.LoadTextStrings
    End If
End Sub

Private Sub Command1_Click()
' Set and save values, then exit form

    Dim u$
    Dim i As Integer
    Dim temp As Single
    Dim accessLevel As Long
    Dim rawCrypt As String

    pressCombo.Tag = "closing"
    update_pressure_unit (pressCombo.ListIndex + 1)
    pressCombo.Tag = ""
    
    update_linear_unit (lengthCombo.ListIndex + 1)
    update_thick_unit (thickCombo.ListIndex + 1)
    update_dens_unit (densityCombo.ListIndex + 1)
    update_mass_unit (massCombo.ListIndex + 1)
    
    
    'added 11-26-07 --Denis
    'save the auto wet feature values
    'WPPS Curr_U$, "permeabilityLoggingFile", permeabilityLoggingFile, IFile$
    ' don't bother with this if the hardware feature is not enabled
    If auto_wet_enable Then
        If auto_wet_check.value = 1 Then
            auto_wet_enable = True
            WPPS Curr_U$, "auto_wet_enable", "Y", IFile$
            auto_wet_used = True
            WPPS Curr_U$, "auto_wet_used", "Y", IFile$
            auto_wet_wet_time = myVal(time1.Text)
            WPPS Curr_U$, "wet_time", str$(auto_wet_wet_time), IFile$
            auto_wet_volume = myVal(txtAutoWetVolume.Text)
            WPPS Curr_U$, "wet_volume", str$(auto_wet_volume), IFile$
            auto_wet_soak_time = myVal(time2.Text)
            WPPS Curr_U$, "soak_time", str$(auto_wet_soak_time), IFile$
            auto_wet_drain_time = myVal(time3.Text)
            WPPS Curr_U$, "drain_time", str$(auto_wet_drain_time), IFile$
            auto_wet_reverse_time = myVal(time4.Text)
            WPPS Curr_U$, "wet_reverse_time", str$(auto_wet_reverse_time), IFile$
            auto_wet_pump_speed = myVal(txtPumpSpeed.Text)
            WPPS Curr_U$, "wet_pump_speed", str$(auto_wet_pump_speed), IFile$
        Else
            auto_wet_used = False
            WPPS Curr_U$, "auto_wet_used", "N", IFile$
        End If
        
        If optWetVolume.value = True Then
            use_auto_wet_volume = True
            WPPS Curr_U$, "use_auto_wet_volume", "Y", IFile$
        Else
            use_auto_wet_volume = False
            WPPS Curr_U$, "use_auto_wet_volume", "N", IFile$
        End If
    End If
    
    'Change comm ports and save results in PorStuff.ini file
    If commPortCombo.ListIndex = 0 Then
        PA = 0
    Else
        PA = val(Right$(commPortCombo.Text, Len(commPortCombo.Text) - 5))
    End If
    WPPS "board.loc", "PA", str$(PA), CSFile$
    
    ' Save log file name if logging
    uselog = (logCheck.value = vbChecked)
    WPPS Curr_U$, "uselog", IIf(uselog, "1", "0"), IFile$
    If uselog Then
        logpath = logText.Caption
        WPPS Curr_U$, "logpath", logpath, IFile$
    End If
    
    permeabilityLogging = (permeabilityLoggingCheckBox.value = vbChecked)
    WPPS Curr_U$, "permeabilityLogging", IIf(permeabilityLogging, "1", "0"), IFile$
    If permeabilityLogging Then
        permeabilityLoggingFile = permeabilityLoggingFileLabel.Caption
        WPPS Curr_U$, "permeabilityLoggingFile", permeabilityLoggingFile, IFile$
    End If
        
    ' Pressure/pore diameter?
    minmaxunits = IIf(endpointoption(3).value, "p", "d")
    WPPS Curr_U$, "minmaxunits", minmaxunits, IFile$
    
    ' Advanced settings only
    advanced_settings = (advancedCheck.value = vbChecked)
    auto_advanced = advanced_settings
    If auto_advanced Then
        WPPS Curr_U$, "auto_advanced", "1", IFile$
    Else
        WPPS Curr_U$, "auto_advanced", "0", IFile$
    End If
    
    ' Second regulator only
    use_second_regulator_only = (regulatorCheck.value = vbChecked)
    If use_second_regulator_only Then 'use second regulator only
        WPPS "capstuff", "use_second_regulator_only", "1", CSFile$
        'WPPS "capstuff", "reg1pmax", "1", CSFile$
        'WPPS "capstuff", "second_regulator_starting_point", "0", CSFile$
    Else
        WPPS "capstuff", "use_second_regulator_only", "0", CSFile$
        'WPPS "capstuff", "reg1pmax", "0", CSFile$
        'WPPS "capstuff", "second_regulator_starting_point", "180", CSFile$
    End If
    
    ' Autoincrement
    auto_increment = (autoincCheck.value = vbChecked)
    If auto_increment Then
        WPPS Curr_U$, "auto_increment", "1", IFile$
    Else
        WPPS Curr_U$, "auto_increment", "0", IFile$
    End If
    '============== New Code edc 05-04-07 =============================================
    'AutoSamplID = (AutoSamplIDcheck.Value = vbChecked)
    'If AutoSamplID Then
     '     WPPS Curr_U$, "auto_sampl_id", "1", IFile$
    'Else
     '   WPPS Curr_U$, "auto_sample_id", "0", IFile$
    'End If
    '============== End New Code ======================================================
    ' Hide sample load prompts
    preloaded_sample = (hidePromptsCheck.value = vbChecked)
    If preloaded_sample Then
        WPPS Curr_U$, "preloaded_sample", "1", IFile$
    Else
        WPPS Curr_U$, "preloaded_sample", "0", IFile$
    End If
    
    ' Use min pressure in dry curve
    use_min_pressure_in_dry = (minPressCheck.value = vbChecked)
    WPPS Curr_U$, "use_min_pressure_in_dry", Format$(IIf(use_min_pressure_in_dry, 1, 0)), IFile$
    
    ' Use min flow in dry curve
    use_min_flow_in_dry = (minFlowCheck.value = vbChecked)
    WPPS Curr_U$, "use_min_flow_in_dry", Format$(IIf(use_min_flow_in_dry, 1, 0)), IFile$
    min_flow_in_dry = myVal(minFlowText.Text) * 1000
    WPPS Curr_U$, "min_flow_in_dry", str$(min_flow_in_dry), IFile$
            
    ' Curve fitting
    autoCurveFit = (curveFitCheck.value = vbChecked)
    WPPS Curr_U$, "auto_curve_fit", IIf(autoCurveFit, "Y", "N"), IFile$
    
    ' auto report type
    For i = 0 To 3
        If resultOption(i).value Then auto_report_type = i
    Next i
    WPPS Curr_U$, "auto_report_type", str$(auto_report_type), IFile$
        
    ' "calibrate" values
    lohmStartMultiplier = val(lohmPercent.Text) / 100
    WPPS Curr_U$, "lohm_start_multiplier", str$(lohmStartMultiplier), IFile$
    WPPS Curr_U$, "lohm_timeout", lohmTimeoutText.Text, IFile$
    WPPS Curr_U$, "lohm_allowable_flow_increase", lohmFlowText.Text, IFile$
    WPPS Curr_U$, "lohm_regulator_increase_factor", lohmRegulatorText.Text, IFile$
    WPPS Curr_U$, "lohm_tolerance", lohmToleranceText.Text, IFile$
    
    ' press hold test options
    pressHoldUnit = IIf(pressHoldUnits(0).value, "S", "M")       ' multiply PSI/sec by 60 to get PSI/min
    WPPS Curr_U$, "pressHoldUnit", pressHoldUnit, IFile$
    num_PH_AvePoints = myVal(numAvePoints.Text)
    WPPS Curr_U$, "num_PH_AvePoints", str$(num_PH_AvePoints), IFile$
    PH_reading_freq = myVal(pholdFreq.Text)
    WPPS Curr_U$, "PH_reading_freq", PH_reading_freq, IFile$
    PH_fail_method$ = IIf(failDPDT(0).value, "dpdt", "dp")
    WPPS Curr_U$, "PH_fail_method", PH_fail_method$, IFile$
    If current_unit% = 1 Then u$ = "" Else u$ = Format$(current_unit%)
    Hold_Rate(current_unit%) = myVal(holdRate.Text)
    WPPS Curr_U$, "Hold_Rate" & u$, str$(Hold_Rate(current_unit%)), IFile$
    Hold_Time(current_unit%) = myVal(testLength.Text)
    WPPS Curr_U$, "Hold_Time" & u$, str$(Hold_Time(current_unit%)), IFile$
    Hold_Delay(current_unit%) = myVal(delayTime.Text)
    WPPS Curr_U$, "Hold_Delay" & u$, str$(Hold_Delay(current_unit%)), IFile$
    PH_stopOnFail = (Check1.value = vbChecked)
    WPPS Curr_U$, "PH_stoponfail", IIf(PH_stopOnFail, "1", "0"), IFile$
    PH_autoscale = (autoscaleCheck.value = vbChecked)
    WPPS Curr_U$, "PH_autoscale", IIf(PH_autoscale, "Y", "N"), IFile$
    PH_minY = myVal(minY.Text)
    PH_maxY = myVal(maxY.Text)
    If PH_minY > PH_maxY Then
        MsgBox ts$(10)          ' ("Minimum Y value for pressure hold test must be smaller than maximum Y value.")
        Exit Sub
    ElseIf PH_maxY = PH_minY Then
        MsgBox ts$(11)          ' ("Maximum and minimum Y values for pressure hold test must be different.")
        Exit Sub
    End If
    WPPS Curr_U$, "PH_minY", str$(PH_minY), IFile$
    WPPS Curr_U$, "PH_maxY", str$(PH_maxY), IFile$
    LP_mintime = myVal(lpMintime.Text)
    If LP_mintime < 0.1 Then LP_mintime = 0.1
    WPPS Curr_U$, "LP_mintime", str$(LP_mintime), IFile$
    PH_regression = regress(1).value
    WPPS Curr_U$, "PH_regression", IIf(PH_regression, "1", "0"), IFile$
    microflowregulator = (microflowregulatorcheck.value = vbChecked)
    WPPS Curr_U$, "microflowregulator", IIf(microflowregulator, "Y", "N"), IFile$
    norefill = (norefillcheck.value = vbChecked)
    WPPS Curr_U$, "norefill", IIf(norefill, "Y", "N"), IFile$
    LP_FlushBeforeTest = (chkLPFlushBeforeTest.value = vbChecked)
    WPPS Curr_U$, "LP_FlushBeforeTest", IIf(LP_FlushBeforeTest, "Y", "N"), IFile$
    LP_CCsToFlush = myVal(txtLPFlushCCs.Text)
    WPPS Curr_U$, "LP_CCsToFlush", str$(LP_CCsToFlush), IFile$
    LP_FlushPressure = myVal(txtLPFlushPressure.Text)
    WPPS Curr_U$, "LP_FlushPressure", str$(LP_FlushPressure), IFile$
    
    lperm_v6_reginccount = val(txtV6LPRegIncAmount.Text)
    WPPS Curr_U$, "LP_V6_RegIncCount", str$(lperm_v6_reginccount), IFile$
    lperm_v6_regincwait = val(txtV6LPRegIncWait.Text)
    WPPS Curr_U$, "LP_V6_RegIncWait", str$(lperm_v6_regincwait), IFile$
    
    LP_DrainAfterTest = (chkLPDrainAfterTest.value = vbChecked)
    WPPS Curr_U$, "LP_DrainAFterTest", IIf(LP_DrainAfterTest, "Y", "N"), IFile$
    LP_DrainTime = myVal(txtLpDrainTime.Text)
    WPPS Curr_U$, "LP_DrainTime", str$(LP_DrainTime), IFile$
    
    If (autocompress Or autopiston) And H2OPERM And (PEN20500 < 0) Then
        delaycompressionliquid = (delaycompressionliquidcheck.value = vbChecked)
        WPPS Curr_U$, "delaycompressionliquid", IIf(delaycompressionliquid, "Y", "N"), IFile$
    End If
    
    ' Microflow test options
    MF_linearSeal = (linSealCheck.value = vbChecked)
    WPPS Curr_U$, "mf_linear_seal", IIf(MF_linearSeal, "Y", "N"), IFile$
    MF_Settle = (mf_settle_check.value = vbChecked)
    WPPS Curr_U$, "mf_settle", IIf(MF_Settle, "Y", "N"), IFile$
    MF_Settle_pressure = myVal(mf_settle_pressure_text.Text)
    WPPS Curr_U$, "mf_settle_pressure", str$(MF_Settle_pressure), IFile$
    MF_Settle_time = myVal(mf_settle_time_text.Text)
    WPPS Curr_U$, "mf_settle_time", str$(MF_Settle_time), IFile$
    MF_recordTemperature = (mf_temperature_check.value = vbChecked)
    WPPS Curr_U$, "mf_record_temperature", IIf(MF_recordTemperature, "Y", "N"), IFile$
    
    ' Curve fit test options
    curve_perc = myVal(cfPercentError.Text)
    If (curve_perc < 0.0001) Then       ' User entered a value that is too small
        curve_perc = 0.0001
        MsgBox ("You entered a percentage error for curve fitting that was too small." + vbCrLf + "Percentage error was reset to 0.0001%")
    ElseIf (curve_perc > 100) Then              ' Value entered is too large
        curve_perc = 100
        MsgBox ("You entered a percentage error for curve fitting that was too large." + vbCrLf + "Percentage error was reset to 100%")
    End If

    curve_nump = Int(myVal(cfNumPoints.Text))
    If curve_nump < 3 Then
        curve_nump = 3
        MsgBox ("You entered a value for the number of curve fit points that was too small." + vbCrLf + "The value was reset to 3 points.")
    ElseIf curve_nump > 10 Then
        curve_nump = 10
        MsgBox ("You entered a value for the number of curve fit points that was too large." + vbCrLf + "The value was reset to 10 points.")
    End If

    curve_maxd = myVal(cfMaxPSI.Text) / PCNV
    If (curve_maxd < 0.0001 / PCNV) Then
        curve_maxd = 0.0001
        MsgBox ("You entered a value for the curve fit maximum pressure difference that was too small." + vbCrLf + "The value was reset to 0.0001")
    End If
    
    WPPS Curr_U$, "curve_maxd", str$(curve_maxd), IFile$
    WPPS Curr_U$, "curve_nump", str$(curve_nump), IFile$
    WPPS Curr_U$, "curve_perc", str$(curve_perc), IFile$
    
    ' GP options
    GP_target = myVal(Text1.Text) / PCNV
    If GP_target <= 0 Then GP_target = 0.01
    GP_delay = myVal(Text2.Text)
    If GP_delay < 0 Then GP_delay = 0
    GP_duration = myVal(Text3.Text)
    If GP_duration < 1 Then GP_duration = 1
    GP_interval = myVal(Text4.Text)
    If GP_interval < 1 Then GP_interval = 1
    GP_numavg = myVal(Text5.Text)
    If GP_numavg < 1 Then GP_numavg = 1
    GP_multiAverageTest = (GP_avgTestCheck.value = vbChecked)
    gP_numAvgTests = myVal(numAvgTests.Text)
    If gP_numAvgTests < 1 Then gP_numAvgTests = 1
    '========edc 09-19-07 set the gas perm multiset test values
    'GPM_targetVol = myVal(txtTargVol.Text)
    'GPM_numberDataPts = CInt(myVal(txtNumberdp.Text))
    'GPM_numberSets = CInt(Val(txtNumSets.Text))
    'GPM_setDuration = CInt(myVal(txtSetDura))
    
    WPPS Curr_U$, "gp_target", str$(GP_target), IFile$
    WPPS Curr_U$, "gp_delay", str$(GP_delay), IFile$
    WPPS Curr_U$, "gp_duration", str$(GP_duration), IFile$
    WPPS Curr_U$, "gp_interval", str$(GP_interval), IFile$
    WPPS Curr_U$, "gp_numavg", str$(GP_numavg), IFile$
    WPPS Curr_U$, "gp_multiavgtest", IIf(GP_multiAverageTest, "Y", "N"), IFile$
    WPPS Curr_U$, "gp_numavgtests", str$(gP_numAvgTests), IFile$
    'gasperm multiset tes values read from the capuser.ini file  edc 09-19-07
    'WPPS Curr_U$, "gp_multisetTest", IIf(GP_multisetTest, "Y", "N"), IFile$
    'WPPS Curr_U$, "gpm_targetVol", Str$(GPM_targetVol), IFile$
    'WPPS Curr_U$, "gpm_numberSets", Str(GPM_numberSets), IFile$
    'WPPS Curr_U$, "gpm_setDuration", Str(GPM_setDuration), IFile$
    'WPPS Curr_U$, "gpm_numberDataPts", Str(GPM_numberDataPts), IFile$
    ' Special options
    zeroTempAtEndOfTest = (zeroTempCheck.value = vbChecked)
    WPPS Curr_U$, "zerotemp", IIf(zeroTempAtEndOfTest, "Y", "N"), IFile$
    BPTLEnable = (BPTLCheck.value = vbChecked)
    BPTLInterval = myVal(BPTLSecondsText.Text) * 10
    BPTLMaxPoints = myVal(BPTLMaxText.Text)
    If BPTLMaxPoints < 1 Then BPTLMaxPoints = 1
    If BPTLInterval < 1 Then BPTLInterval = 1
    WPPS Curr_U$, "BPTLEnable", IIf(BPTLEnable, "Y", "N"), IFile$
    WPPS Curr_U$, "BPTLInterval", str$(BPTLInterval / 10#), IFile$
    WPPS Curr_U$, "BPTLMaxPoints", str$(BPTLMaxPoints), IFile$
    bubbler_selected = (BubblerCheck.value = vbChecked)
    WPPS Curr_U$, "bubbler_selected", IIf(bubbler_selected, "Y", "N"), IFile$
    'edc 03-05-07 Language choice support
       For i = 1 To languageCount
        If available_languages$(i, 2) = Combo1.Text Then
            currentLanguagePath$ = available_languages(i, 1)
        End If
    Next i
    
    'AJB 10-27-09
    If testGasCheck.value = vbChecked Then
        testGasStatus = True
    Else
        testGasStatus = False
    End If
    
    WPPS Curr_U$, "testGasStatus", IIf(testGasStatus, "Y", "N"), IFile$
    
    If depressurizeCheck.value = vbChecked Then
        depressurizeBeforeTest = True
    Else
        depressurizeBeforeTest = False
    End If
    WPPS Curr_U$, "DepressurizeBeforeTest", IIf(depressurizeBeforeTest, "Y", "N"), IFile$
    
    If savePreBPdataCheck.value = vbChecked Then
        savePreBPdata = True
    Else
        savePreBPdata = False
    End If
    WPPS Curr_U$, "SavePreBPdata", IIf(savePreBPdata, "Y", "N"), IFile$
    
    'JF 2-11-10
    'Save changes for Additional Information
    useAdditionalInfo = IIf(chkEnableAdditionalInfo = 1, True, False)
    numberOfAdditionalInfoLines = val(txtNumberOfInfoLines.Text)
    ReDim infoLineHeaders(numberOfAdditionalInfoLines)
    For i = 0 To numberOfAdditionalInfoLines - 1
        lstInfoLines.ListIndex = i
        infoLineHeaders(i) = lstInfoLines.Text
    Next i
    
    WPPS Curr_U$, "UseAdditionalInfo", IIf(useAdditionalInfo, "Y", "N"), IFile$
    WPPS Curr_U$, "NumberOfAdditionalInfoLines", str$(numberOfAdditionalInfoLines), IFile$
    For i = 0 To numberOfAdditionalInfoLines - 1
        WPPS Curr_U$, "InfoLineHeader" & i, infoLineHeaders(i), IFile$
    Next
    
    'JF 02-16-2010
    'Save changes for Humidity Controls
    recordHumidityForAutoTests = IIf(chkRecordHumidityForAutoTests.value = 1, True, False)
    enableHumidityControlForAutoTests = IIf(chkEnableHumidityControl.value = 1, True, False)
    targetHumidity = val(txtTargetHumidity.Text)
    goToHumidityMaxWaitTime = val(txtMaxAdjustmentWaitTime.Text)
    goToHumidityMinWaitTime = val(txtMinAdjustmentWaitTime.Text)
    goToHumidityTolerance = val(txtTargetTolerance.Text)
    stableHumidityMaxWaitTime = val(txtMaxStabilityWaitTime.Text)
    stableHumidityMinWaitTime = val(txtMinStabilityWaitTime.Text)
    stableHumidityTolerance = val(txtStabilityTolerance.Text)
    stableHumiditySleepTime = val(txtStabilitySleepTime.Text)
    initialHumidityWaitTime = val(txtInitialHumidityWaitTime.Text)
    minHumidityAdjustmentFlow = val(txtMininmumAdjustmentFlow.Text)
    
    WPPS Curr_U$, "RecordHumidityForAutoTests", chkRecordHumidityForAutoTests.value, IFile$
    WPPS Curr_U$, "EnableHumidityControlForAutoTests", chkEnableHumidityControl.value, IFile$
    WPPS Curr_U$, "TargetHumidity", str$(targetHumidity), IFile$
    WPPS Curr_U$, "GoToHumidityMaxWaitTime", str$(goToHumidityMaxWaitTime), IFile$
    WPPS Curr_U$, "GoToHumidityMinWaitTime", str$(goToHumidityMinWaitTime), IFile$
    WPPS Curr_U$, "GoToHumidityTolerance", str$(goToHumidityTolerance), IFile$
    WPPS Curr_U$, "StableHumidityMaxWaitTime", str$(stableHumidityMaxWaitTime), IFile$
    WPPS Curr_U$, "StableHumidityMinWaitTime", str$(stableHumidityMinWaitTime), IFile$
    WPPS Curr_U$, "StableHumidityTolerance", str$(stableHumidityTolerance), IFile$
    WPPS Curr_U$, "StableHumiditySleepTime", str$(stableHumiditySleepTime), IFile$
    WPPS Curr_U$, "InitialHumidityWaitTime", str$(initialHumidityWaitTime), IFile$
    WPPS Curr_U$, "MinHumidityAdjustmentFlow", str$(minHumidityAdjustmentFlow), IFile$
    
    i = val(txtBPPointDetectionCount.Text)
    If i > 0 Then
        BP_PointDetectionCount = i
        WPPS Curr_U$, "BP_PointDetectionCount", txtBPPointDetectionCount.Text, IFile$
    End If
    
    If UAC_Changed Then
        If UAC_Enabled Then
            WPPS "UAC", "UAC_Enabled", "Y", UACFile
        Else
            WPPS "UAC", "UAC_Enabled", "N", UACFile
        End If
        
        WPPS "UAC", "UAC_UserCount", Trim$(str$(UAC_UserCount)), UACFile
        For i = 1 To UAC_UserCount
            WPPS "UAC", "UAC_Username" + Trim$(str$(i)), UAC_Users(i).Username, UACFile
            'rawCrypt = UAC_Crypt.URLEncodeBinaryData(UAC_Users(i).Password)
            WPPS "UAC", "UAC_Password" + Trim$(str$(i)), rawCrypt, UACFile
            accessLevel = accessLevel + (2 ^ (i - 1) * UAC_Users(i).accessLevel)
        Next i
        WPPS "UAC", "UAC_AL", Trim$(str$(accessLevel)), UACFile
    End If
    
   ' reset originalLanguageIndex so the form unload doesn't set the language back
    '  (the form unload can set the language back in case you close the form without saving)
    originalLanguageIndex = languageIndex
    
   WPPS "default", "current_language_path", currentLanguagePath$, IFile$
    CAPFLOW.LoadTextStrings
    TitleScrn.LoadTextStrings
    ' Get out!
    Unload prefsForm
    
End Sub

Private Sub Command2_Click()
    Unload prefsForm
End Sub



Private Sub Command3_Click()
Load BPSettings
BPSettings.Show (1)
End Sub



Private Sub Form_Load()
    
    Dim i As Integer
    Dim cc As COMMCONFIG, L As Long
    Dim ports(25, 2) As Integer, portcount As Integer
    
    LoadTextStrings
    SSTab1.TabEnabled(9) = False        ' TEMP UNTIL WINDOW IS FINISHED

    UAC_Changed = False
    
    If Not supervisor Then
        SSTab1.TabEnabled(14) = False
    Else
        For i = 1 To UAC_UserCount
            lstUsers.AddItem UAC_Users(i).Username
        Next i
        If UAC_Enabled Then chkEnableUAC.value = 1
        If lstUsers.ListCount > 0 Then
            lstUsers.ListIndex = 0
        End If
    End If
    ' Need to load in all the options and set them to their current values:
    
    
    ' Units:
    If liqpermonly = True Then
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(5) = False
        SSTab1.TabEnabled(7) = False
    End If
    pressCombo.Tag = "init"
    BuildPressureList
    pressCombo.Tag = ""

    With lengthCombo
        .clear
        .AddItem ("mm")
        .AddItem ("cm")
        .AddItem ("inch")
        .ListIndex = linear_unit_index% - 1
    End With

    With thickCombo
        .clear
        .AddItem ("mm")
        .AddItem ("cm")
        .AddItem ("mil")
        .ListIndex = thick_unit_index% - 1
    End With
    
    With densityCombo
        .clear
        .AddItem ("g/cm^3")
        .AddItem ("lb/in^3")
        .ListIndex = dens_unit_index% - 1
    End With
    
    With massCombo
        .clear
        .AddItem ("g")
        .AddItem ("lb")
        .ListIndex = mass_unit_index% - 1
    End With
    
    ' Comm port setup
    With commPortCombo
        .clear
        .AddItem (ts$(13))      ' "Demo Mode"
        portcount = 0
        For i% = 1 To 25
            L = LenB(cc)
            If GetDefaultCommConfig("COM" + Format$(i), cc, L) Then
                .AddItem (ts$(14) + " " + str$(i))     ' "Comm"
                portcount = portcount + 1
                ports(portcount, 1) = portcount
                ports(portcount, 2) = i
            End If
        Next i
        
        If PA < 1 Then
            .ListIndex = 0
        Else
        ' Roundabout way of doing things, but seemingly necessary because ports
        ' may not be consecutive
            For i = 1 To portcount
                If ports(i, 2) = PA Then .ListIndex = ports(i, 1)
            Next i
        End If
    End With
    
    ' Logging options
    logCheck.value = IIf(uselog, vbChecked, vbUnchecked)
    logText.Caption = logpath
    logText.Visible = uselog
    Label6.Visible = uselog
    changeFileButton.Visible = uselog
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Added 11-26-07 --Denis
    'Auto Wet Feature tab settings
    'Dim wetVal, soakVal, drainVal As String     'create the 3 time variables
    'If test is not hardware supported then set all value to 0 or OFF or defult values
    ' 12-12-07 - rvw - now simply disable the tab if you don't use it
    If Not auto_wet_enable Then
        SSTab1.TabEnabled(10) = False
        test_unsupport.Visible = True
        auto_wet_check.Visible = False
'     'Set wet test values to 0 since test is not supported
        WPPS "Capstuff", "auto_wet_used", "0", CSFile$
'     '1st option
        optWetTime.Visible = False
        optWetVolume.Visible = False
        txtAutoWetVolume.Visible = False
        Label26.Visible = False
        wet_time.Visible = False
        time1.Visible = False
        Label27.Visible = False
        time4.Visible = False
        time4Label.Visible = False
        txtPumpSpeed.Visible = False
        lblPumpSpeed.Visible = False
        lblPumpSpeedDesc.Visible = False
        
'     '2nd option
        soak_time.Visible = False
        time2.Visible = False
        Label28.Visible = False
'     '3rd option
        drain_time.Visible = False
        time3.Visible = False
        Label29.Visible = False
'     'Set values to 0 since test is not supported
        WPPS Curr_U$, "wet_time", "0", IFile$
        WPPS Curr_U$, "soak_time", "0", IFile$
        WPPS Curr_U$, "drain_time", "0", IFile$
    Else
        If auto_wet_used = True Then
            auto_wet_check.value = 1 ' default it to 1 so that if we set it to 0 later it will activate the handler to update the GUI
            If auto_wet_enable = False Then auto_wet_check.value = 0 'DUMB!
            ' fill in all the time values, even if we are disabled, so that when they enable they will see the proper values
            time1.Text = Format$(auto_wet_wet_time)
            txtAutoWetVolume.Text = Format$(auto_wet_volume)
            time2.Text = Format$(auto_wet_soak_time)
            time3.Text = Format$(auto_wet_drain_time)
            time4.Text = Format$(auto_wet_reverse_time)
            txtPumpSpeed.Text = Format$(auto_wet_pump_speed)
    '        test_unsupport.Visible = False
            
            If use_auto_wet_volume Then optWetVolume.value = True
        Else
            auto_wet_check.value = 0
            '1st option
            optWetTime.Visible = False
            optWetVolume.Visible = False
            txtAutoWetVolume.Visible = False
            Label26.Visible = False
            wet_time.Visible = False
            time1.Visible = False
            Label27.Visible = False
            time4.Visible = False
            time4Label.Visible = False
            txtPumpSpeed.Visible = False
            lblPumpSpeed.Visible = False
            lblPumpSpeedDesc.Visible = False
            
         '2nd option
            soak_time.Visible = False
            time2.Visible = False
            Label28.Visible = False
         '3rd option
            drain_time.Visible = False
            time3.Visible = False
            Label29.Visible = False
        End If
    End If
    
'added 11-21-07 by Denis for auto wet test
'    auto_wet_check.value = Val(gpps2("Capstuff", "auto_wet_used", CSFile$, "0"))     'Unchecked  DEFUALT
'    '"Auto wet test" tab options
'    If gpps2("Capstuff", "auto_wet_used", CSFile$, "0") = "0" Then        'Hide the options if test is off.
'     '1st option
'        wet_time.Visible = False
'        time1.Visible = False
'        Label27.Visible = False
'     '2nd option
'        soak_time.Visible = False
'        time2.Visible = False
'        Label28.Visible = False
'     '3rd option
'        drain_time.Visible = False
'        time3.Visible = False
'        Label29.Visible = False
'    End If

    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ReDim Preserve TType(30)
    If TType(current_unit%) = 3 Then
        GP_avgTestCheck.Enabled = True
    Else
        GP_avgTestCheck.Enabled = False
    End If
        
    
    ' gas permeability logging options
    permeabilityLoggingCheckBox.value = IIf(permeabilityLogging, vbChecked, vbUnchecked)
    permeabilityLoggingFileLabel.Caption = permeabilityLoggingFile
    permeabilityLoggingFileLabel.Visible = permeabilityLogging
    Label16.Visible = permeabilityLogging
    permeabilityLoggingSelectButton.Visible = permeabilityLogging
    
    ' pressure/pore diameter
    endpointoption(2).value = (minmaxunits = "d")
    endpointoption(3).value = (minmaxunits = "p")
    
    ' "Test" tab checkboxes
    advancedCheck.value = IIf(advanced_settings, vbChecked, vbUnchecked)
    regulatorCheck.value = IIf(use_second_regulator_only, vbChecked, vbUnchecked)
    autoincCheck.value = IIf(auto_increment, vbChecked, vbUnchecked)
    'AutoSamplIDcheck.Value = IIf(AutoSamplID, vbChecked, vbUnchecked)
    hidePromptsCheck.value = IIf(preloaded_sample, vbChecked, vbUnchecked)
    minPressCheck.value = IIf(use_min_pressure_in_dry, vbChecked, vbUnchecked)
    minFlowCheck.value = IIf(use_min_flow_in_dry, vbChecked, vbUnchecked)
    minFlowText.Text = Format$(min_flow_in_dry / 1000)
    
    curveFitCheck.value = IIf(autoCurveFit, vbChecked, vbUnchecked)
    '*******
    'we are hiding "show results at end of test" at KG's behest; results from Capwin and
    'caprep are not consistent
    '6/1/05: This is no longer necessary! The results match!
    'resultOption(3).Visible = False
   ' If auto_report_type = 3 Then auto_report_type = 0
    resultOption(auto_report_type).value = True
    '*******
    
    ' "calibrate" options
    lohmPercent.Tag = "init"
    lohmPercent.Text = Format$(lohmStartMultiplier * 100)
    lohmPercent.Tag = ""
    lohmTimeoutText.Text = myVal(gpps2(Curr_U$, "lohm_timeout", IFile$, "50"))
    lohmFlowText.Text = myVal(gpps2(Curr_U$, "lohm_allowable_flow_increase", IFile$, "1000"))
    lohmRegulatorText.Text = myVal(gpps2(Curr_U$, "lohm_regulator_increase_factor", IFile$, "1"))
    lohmToleranceText.Text = myVal(gpps2(Curr_U$, "lohm_tolerance", IFile$, "50"))
  ReDim Preserve Hold_Time(10)
  ReDim Preserve Hold_Delay(10)
  ReDim Preserve Hold_Rate(10)
    ' press hold test options
    If pressHoldUnit = "S" Then pressHoldUnits(0).value = True Else pressHoldUnits(1).value = True
    numAvePoints.Text = Format$(num_PH_AvePoints)
    pholdFreq.Text = Format$(PH_reading_freq)
    failDPDT(0).value = (PH_fail_method$ = "dpdt")
    failDPDT(1).value = (PH_fail_method$ = "dp")
    Hold_Time(current_unit%) = val(gpps2(Curr_U$, "Hold_Time", IFile$, "10"))
    Hold_Delay(current_unit%) = val(gpps2(Curr_U$, "Hold_Delay", IFile$, "0"))
    Hold_Rate(current_unit%) = val(gpps2(Curr_U$, "Hold_Rate", IFile$, "0"))
    testLength.Text = Format$(Hold_Time(current_unit%))
    delayTime.Text = Format$(Hold_Delay(current_unit%))
    holdRate.Text = Format$(Hold_Rate(current_unit%))
    Check1.value = IIf(PH_stopOnFail, vbChecked, vbUnchecked)
    autoscaleCheck.value = IIf(PH_autoscale, vbChecked, vbUnchecked)
    Frame8.Visible = (autoscaleCheck.value = vbUnchecked)
    minY.Text = Format$(PH_minY)
    maxY.Text = Format$(PH_maxY)
    If PH_regression Then regress(1).value = True Else regress(0).value = True
    
    ' If not checking values during PH test, how can we abort if the test fails? Leave
    ' hidden for now
    Check1.Visible = False
    
    
    ' Liquid perm options
    lpMintime.Text = Format$(LP_mintime)
    norefillcheck.value = IIf(norefill, vbChecked, vbUnchecked)
    chkLPFlushBeforeTest.value = IIf(LP_FlushBeforeTest, vbChecked, vbUnchecked)
    txtLPFlushCCs.Text = Format$(LP_CCsToFlush)
    txtLPFlushPressure.Text = Format$(LP_FlushPressure)
    
    If version < 7 Then
        txtV6LPRegIncAmount.Visible = True
        lblV6LPRegIncAmount.Visible = True
        txtV6LPRegIncWait.Visible = True
        lblV6LPRegIncWait.Visible = True
        txtV6LPRegIncAmount.Text = lperm_v6_reginccount
        txtV6LPRegIncWait.Text = lperm_v6_regincwait
    Else
        txtV6LPRegIncAmount.Visible = False
        lblV6LPRegIncAmount.Visible = False
        txtV6LPRegIncWait.Visible = False
        lblV6LPRegIncWait.Visible = False
    End If
    
    chkLPDrainAfterTest.value = IIf(LP_DrainAfterTest, vbChecked, vbUnchecked)
    txtLpDrainTime.Text = Format$(LP_DrainTime)
    
    If (autocompress Or autopiston) And H2OPERM And (PEN20500 < 0) Then
        delaycompressionliquidcheck.value = IIf(delaycompressionliquid, vbChecked, vbUnchecked)
    Else
        delaycompressionliquidcheck.Visible = False
    End If
    
    ' microflow options
    microflowregulatorcheck.value = IIf(microflowregulator, vbChecked, vbUnchecked)
    linSealCheck.value = IIf(MF_linearSeal, vbChecked, vbUnchecked)
    mf_settle_check.value = IIf(MF_Settle, vbChecked, vbUnchecked)
    mf_settle_pressure_text.Text = Format$(MF_Settle_pressure)
    mf_settle_time_text.Text = Format$(MF_Settle_time)
    mf_temperature_check.value = IIf(MF_recordTemperature, vbChecked, vbUnchecked)
    
    ' Curve fit options
    cfPercentError.Text = Format$(curve_perc)
    cfPercentError.ToolTipText = ts$(16)
    cfNumPoints.Text = Format$(curve_nump)
    cfNumPoints.ToolTipText = ts$(15)
    cfMaxPSI.Text = Format$(curve_maxd * PCNV)
    cfMaxPSI.ToolTipText = ts$(17) + " (" + PU$ + ")"
    Label10.Caption = ts$(18) + " " + ts$(19)
    
    ' Gasperm options Single Point Options
    Text1.Text = Format(GP_target * PCNV, "###0.0###")
    Text2.Text = Format$(GP_delay)
    Text3.Text = Format$(GP_duration)
    Text4.Text = Format$(GP_interval)
    Text5.Text = Format$(GP_numavg)
   'edc 10-10-07 removed Multi Set Options
   ' txtTargVol.Text = Format(GPM_targetVol, "###0.0###")
   ' txtNumSets.Text = Format(GPM_numberSets)
   ' txtSetDura.Text = Format(GPM_setDuration)
   ' txtNumberdp.Text = Format(GPM_numberDataPts)
    Label11(5).Caption = Label11(5).Caption + " (" + PU$ + ")"
    numAvgTests.Text = str$(gP_numAvgTests)
    GP_avgTestCheck.value = IIf(GP_multiAverageTest, vbChecked, vbUnchecked)
    
    ' Special options
    zeroTempCheck.value = IIf(zeroTempAtEndOfTest, vbChecked, vbUnchecked)
    BPTLCheck.value = IIf(BPTLEnable, vbChecked, vbUnchecked)
    BPTLSecondsText.Text = Format$(BPTLInterval / 10#)
    BPTLMaxText.Text = Format$(BPTLMaxPoints)
    If bubbler_enable Then
        BubblerCheck.value = IIf(bubbler_selected, vbChecked, vbUnchecked)
    Else
        BubblerCheck.Visible = False
    End If
 
    If doorlock Then
        Text6.Visible = True
        Label25.Visible = True
        Text6.Text = safe_temperature
    Else
        Text6.Visible = False
        Label25.Visible = False
        
    End If
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    'edc 03-02-07 Language Box
    Combo1.clear
    currentLanguagePath = gpps2("default", "current_language_path", IFile$, "CapWinLanguageEN.ini")
    For i = 1 To languageCount
        Combo1.AddItem (available_languages(i, 2))
        If available_languages$(i, 1) = currentLanguagePath$ Then
            Combo1.ListIndex = i - 1
            languageIndex = i - 1
        End If
    Next i
    
    'AJB 11-02-09
    testGasCheck.value = IIf(testGasStatus, vbChecked, vbUnchecked)
    
    depressurizeCheck.value = IIf(depressurizeBeforeTest, vbChecked, vbUnchecked)
    
    savePreBPdataCheck.value = IIf(savePreBPdata, vbChecked, vbUnchecked)
    
    'JF 2-11-10
    If useAdditionalInfo Then
        chkEnableAdditionalInfo.value = 1
    Else
        chkEnableAdditionalInfo.value = 0
    End If
    
    lblNumOfInfoLines.Visible = useAdditionalInfo
    txtNumberOfInfoLines.Visible = useAdditionalInfo
    lblInfoHeaders.Visible = useAdditionalInfo
    lstInfoLines.Visible = useAdditionalInfo
    
    txtNumberOfInfoLines.Text = numberOfAdditionalInfoLines
    For i = 0 To numberOfAdditionalInfoLines - 1
        lstInfoLines.AddItem infoLineHeaders(i)
    Next i
    
    'JF 02-16-2010
    If hasHumidityControls Then
        SSTab1.TabEnabled(12) = True
    Else
        SSTab1.TabEnabled(12) = False
    End If
    
    If recordHumidityForAutoTests Then
        chkRecordHumidityForAutoTests.value = 1
    Else
        chkRecordHumidityForAutoTests.value = 0
    End If
    
    If enableHumidityControlForAutoTests Then
        chkEnableHumidityControl.value = 1
    Else
        chkEnableHumidityControl.value = 0
    End If
    
    txtTargetHumidity.Text = targetHumidity
    txtMaxAdjustmentWaitTime.Text = goToHumidityMaxWaitTime
    txtMinAdjustmentWaitTime.Text = goToHumidityMinWaitTime
    txtTargetTolerance.Text = goToHumidityTolerance
    txtMaxStabilityWaitTime.Text = stableHumidityMaxWaitTime
    txtMinStabilityWaitTime.Text = stableHumidityMinWaitTime
    txtStabilityTolerance.Text = stableHumidityTolerance
    txtStabilitySleepTime.Text = stableHumiditySleepTime
    txtInitialHumidityWaitTime.Text = initialHumidityWaitTime
    txtMininmumAdjustmentFlow.Text = minHumidityAdjustmentFlow
    
    txtBPPointDetectionCount.Text = BP_PointDetectionCount
    
    Me.LoadTextStrings
    SSTab1.TabVisible(0) = True
    SSTab1.Tab = 0

End Sub

Private Sub Label38_Click()

End Sub

Private Sub logCheck_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Toggle visibility of log details

    logText.Visible = (logCheck.value = vbChecked)
    changeFileButton.Visible = (logCheck.value = vbChecked)
    Label6.Visible = (logCheck.value = vbChecked)
    
End Sub
Private Sub lohmPercent_LostFocus()
' Validate the value entered for lohm starting percentage

    If (val(lohmPercent.Text) > 100 Or val(lohmPercent.Text) <= 0) And (lohmPercent.Tag = "") Then     ' Illegal percentage values
        MsgBox ts$(12)  ' ("Percentage entered must be between 0 and 100")
        lohmPercent.Tag = "fix"
        lohmPercent.Text = ""
        lohmPercent.Tag = ""
    End If
    
End Sub
Private Sub lohmRegulatorText_LostFocus()
    If (val(lohmRegulatorText.Text) > 5 Or val(lohmRegulatorText.Text) <= 0) And (lohmRegulatorText.Tag = "") Then     ' Illegal percentage values
        MsgBox "Value entered must be between 1 and 5"
        lohmRegulatorText.Tag = "fix"
        lohmRegulatorText.Text = ""
        lohmRegulatorText.Tag = ""
    End If
End Sub

Private Sub lstUsers_Click()
    If lstUsers.ListIndex = 0 Then
        optAccountType(0).value = True
        optAccountType(0).Enabled = False
        optAccountType(1).Enabled = False
    Else
        optAccountType(0).Enabled = True
        optAccountType(1).Enabled = True
        If UAC_Users(lstUsers.ListIndex + 1).accessLevel = 0 Then
            optAccountType(0).value = True
        Else
            optAccountType(1).value = True
        End If
    End If
End Sub

Private Sub optAccountType_Click(Index As Integer)
    Dim lstIndex As Integer
    
    If UAC_Enabled Then
        lstIndex = lstUsers.ListIndex + 1
        If lstIndex > 1 Then
            UAC_Users(lstIndex).accessLevel = Index
            UAC_Changed = True
        End If
    End If
End Sub

Private Sub permeabilityLoggingCheckBox_Click()

    permeabilityLoggingFileLabel.Visible = (permeabilityLoggingCheckBox.value = vbChecked)
    permeabilityLoggingSelectButton.Visible = (permeabilityLoggingCheckBox.value = vbChecked)
    Label16.Visible = (permeabilityLoggingCheckBox.value = vbChecked)

End Sub

Private Sub permeabilityLoggingSelectButton_Click()

    fsel_name$ = permeabilityLoggingFileLabel.Caption
    fsel_title$ = ts$(9)                ' "Set Logging File"
    fsel_io = False ' file doesn't have to exist
    fsel_path = "*.txt"
    fsel Me.hwnd
    If fsel_return$ <> "" Then
        permeabilityLoggingFileLabel.Caption = fsel_return$
    End If

End Sub

Private Sub pressCombo_Click()
' Note: with this setup, pressure units get set as soon as user selects new value (unlike others, which wait
' until user clicks "save"

    If pressCombo.Tag <> "autoclick" Then       ' autoclick = event generated by software call to set .listindex
        update_pressure_unit (pressCombo.ListIndex + 1)
    End If
    
End Sub

Private Sub setFontButton_Click()
' Activate the font selection window and save the new font values

    selectfont.Show 1
    WPPS "default", "default_fontname", default_font.font, language$
    WPPS "default", "default_fontsize", str$(default_font.fontsize), language$
    WPPS "default", "default_fontbold", IIf(default_font.fontbold, "Y", "N"), language$
    LoadTextStrings
    TitleScrn.LoadTextStrings

End Sub

Sub update_pressure_unit(Index As Integer)
' Handle clicking on the pressure unit selector in the prefs window.

    Dim temp$, Ret$, i As Integer
    Dim r As Long
    
    If pressCombo.Tag = "init" Then Exit Sub        ' Prevent sub from being called by spurious onClick events
    
    Select Case Index
        Case 1
            PCNV = 1
            PU$ = "PSI"
        Case 2
            PCNV = 6.894733
            PU$ = "KPA"
        Case 3
            PCNV = 0.070307
            PU$ = "Kg/cm2"
        Case 4
            PCNV = 0.06894733
            PU$ = "BAR"
        Case Is > 4
            temp$ = "unitdef" + LTrim$(RTrim$(str$(Index - 4)))
            Ret$ = String$(255, " ")
            r = GPPS("UnitDefs", temp$, "---", Ret$, 255, CSFile$)
            If nulltrim(Ret$) = "---" Then
GetName:        GetValue.Label1.Caption = ts$(1)       ' "Enter Unit Name: (6 char)"
                GetValue.Text1.Text = "ABC"
                GetValue.Label1.Tag = "text"
                GetValue.Continue.default = True
                GetValue.Show 1
                GetValue.Label1.Tag = ""
                If Got_Value <> -9 Then
                    PU$ = Got_Text
                Else
                    Exit Sub
                End If
                If Len(PU$) > 6 Then
                    MsgBox ts$(2), 0, ts$(3)       ' "The unit name must be less than 7 characters."/"New Unit"
                    GoTo GetName
                End If
GetFact:        GetValue.Label1.top = 0
                GetValue.Label1.Height = 615
                GetValue.Label1.Caption = ts$(4) + ":" + vbCrLf + "(" + ts$(5) + vbCrLf + ts$(6) + ".)" '"Enter Conversion Factor"/"This 1 PSI in this new unit."/"i.e. 1 PSI = 0.0689 BAR"
                GetValue.Text1.Text = "1"
                GetValue.Label1.Tag = ""
                GetValue.Continue.default = True
                GetValue.Show 1
                GetValue.Label1.top = 240
                GetValue.Label1.Height = 255
                If Got_Value <> -9 Then
                    PCNV = Got_Value
                Else
                    Exit Sub
                End If
                If Got_Value < 0 Then
                    MsgBox ts$(7), 0, ts$(3)       ' "The unit factor must be positive"/"New Unit"
                    GoTo GetFact
                End If
                temp$ = "unitdef" + LTrim$(RTrim$(str$(Index - 4)))
                WPPS "UnitDefs", temp$, PU$, CSFile$
                WPPS "UnitDefs", temp$ + "num", str$(PCNV), CSFile$
                BuildPressureList
            Else
                PU$ = nulltrim(Ret$)
                Ret$ = String$(255, " ")
                r = GPPS("UnitDefs", temp$ + "num", "1", Ret$, 255, CSFile$)
                PCNV = val(Ret$)
            End If
    End Select
    With pressCombo
        .Tag = "autoclick"
        .ListIndex = Index - 1
        .Tag = ""
    End With
    TitleScrn.Label6.Caption = PU$
    WPPS Curr_U$, "unit", PU$, IFile$
    WPPS Curr_U$, "unitval", str$(PCNV), IFile$

End Sub

Sub BuildPressureList()
' Create a new list of pressure units for the drop-down menu

    With pressCombo
        .clear
        .AddItem ("PSI")
        .AddItem ("KPA")
        .AddItem ("kg/cm2")
        .AddItem ("BAR")
        .AddItem (gpps2("UnitDefs", "unitdef1", CSFile$, "---"))
        .AddItem (gpps2("UnitDefs", "unitdef2", CSFile$, "---"))
        .AddItem (gpps2("UnitDefs", "unitdef3", CSFile$, "---"))
        .AddItem (gpps2("UnitDefs", "unitdef4", CSFile$, "---"))
        update_units_check (PU$)
        .ListIndex = press_unit_index%
    End With

End Sub

Sub LoadTextStrings()
' Load in text from language.ini

    Dim i As Integer

    ' Form elements
    For i = 0 To 8
        SSTab1.TabCaption(i) = gpps2("prefs", "sstab1" + str$(i), language$, SSTab1.TabCaption(i))
    Next i
    set_fontname SSTab1, default_font
    Frame1.Caption = get_thing("prefs", "frame1", language$, Frame1.Caption, Frame1, default_font)
    Frame2.Caption = get_thing("prefs", "frame2", language$, Frame2.Caption, Frame2, default_font)
    Frame3.Caption = get_thing("prefs", "frame3", language$, Frame3.Caption, Frame3, default_font)
    Frame4.Caption = get_thing("prefs", "frame4", language$, Frame4.Caption, Frame4, default_font)
    Frame5.Caption = get_thing("prefs", "frame5", language$, Frame5.Caption, Frame5, default_font)
    Frame7.Caption = get_thing("prefs", "frame7", language$, Frame7.Caption, Frame7, default_font)
    Frame9.Caption = get_thing("prefs", "frame9", language$, Frame9.Caption, Frame9, default_font)
    Frame11.Caption = get_thing("prefs", "frame11", language$, Frame11.Caption, Frame11, default_font)
    Frame12.Caption = get_thing("prefs", "frame12", language$, Frame12.Caption, Frame12, default_font)
    Frame14.Caption = get_thing("prefs", "frame14", language$, Frame14.Caption, Frame14, default_font)
    
    set_fontstuff pressCombo, default_font
    set_fontstuff thickCombo, default_font
    set_fontstuff densityCombo, default_font
    set_fontstuff lengthCombo, default_font
    set_fontstuff massCombo, default_font
    set_fontstuff commPortCombo, default_font
    
    Label1.Caption = get_thing("prefs", "label1", language$, Label1.Caption, Label1, default_font)
    Label2.Caption = get_thing("prefs", "label2", language$, Label2.Caption, Label2, default_font)
    Label3.Caption = get_thing("prefs", "label3", language$, Label3.Caption, Label3, default_font)
    label4.Caption = get_thing("prefs", "label4", language$, label4.Caption, label4, default_font)
    label5.Caption = get_thing("prefs", "label5", language$, label5.Caption, label5, default_font)
    Label8.Caption = get_thing("prefs", "label8", language$, Label8.Caption, Label8, default_font)
    Label9.Caption = get_thing("prefs", "label9", language$, Label9.Caption, Label9, default_font)
    Label13.Caption = get_thing("prefs", "label13", language$, Label13.Caption, Label13, default_font)
    
    logCheck.Caption = get_thing("prefs", "logcheck", language$, logCheck.Caption, logCheck, default_font)
    Label6.Caption = get_thing("prefs", "label6", language$, Label6.Caption, Label6, default_font)
    logText.Caption = get_thing("prefs", "logtext", language$, logText.Caption, logText, default_font)
    set_fontname changeFileButton, default_font
    changeFileButton.Caption = gpps2("prefs", "changefile", language$, changeFileButton.Caption)
    set_fontname setFontButton, default_font
    setFontButton.Caption = gpps2("prefs", "setfont", language$, setFontButton.Caption)
    set_fontname Command1, default_font
    Command1.Caption = gpps2("prefs", "command1", language$, Command1.Caption)
    set_fontname Command2, default_font
    Command2.Caption = gpps2("prefs", "command2", language$, Command2.Caption)
    
    endpointoption(2).Caption = get_thing("prefs", "endpoint2", language$, endpointoption(2).Caption, endpointoption(2), default_font)
    endpointoption(3).Caption = get_thing("prefs", "endpoint3", language$, endpointoption(3).Caption, endpointoption(3), default_font)
    
    advancedCheck.Caption = get_thing("prefs", "advanced", language$, advancedCheck.Caption, advancedCheck, default_font)
    autoincCheck.Caption = get_thing("prefs", "autoinc", language$, autoincCheck.Caption, autoincCheck, default_font)
    hidePromptsCheck.Caption = get_thing("prefs", "hideprompt", language$, hidePromptsCheck.Caption, hidePromptsCheck, default_font)
    minPressCheck.Caption = get_thing("prefs", "minpress", language$, minPressCheck.Caption, minPressCheck, default_font)
    minFlowCheck.Caption = get_thing("prefs", "minflow", language$, minFlowCheck.Caption, minFlowCheck, default_font)
    autoscaleCheck.Caption = get_thing("prefs", "autoscalecheck", language$, autoscaleCheck.Caption, autoscaleCheck, default_font)
    curveFitCheck.Caption = get_thing("prefs", "curvefitcheck", language$, curveFitCheck.Caption, curveFitCheck, default_font)
    
    lohmCaption.Caption = get_thing("prefs", "lohmcaption", language$, lohmCaption.Caption, lohmCaption, default_font)
    lohmCaption2.Caption = get_thing("prefs", "lohmcaption2", language$, lohmCaption2.Caption, lohmCaption2, default_font)
    lohmCaption3.Caption = get_thing("prefs", "lohmcaption3", language$, lohmCaption3.Caption, lohmCaption3, default_font)
    lohmCaption4.Caption = get_thing("prefs", "lohmcaption4", language$, lohmCaption4.Caption, lohmCaption4, default_font)
    set_fontstuff lohmCaption5, default_font
    Label12.Caption = get_thing("prefs", "label12", language$, Label12.Caption, Label12, default_font)
    For i = 0 To 7
        pLabel(i).Caption = get_thing("prefs", "plabel" + Format$(i), language$, pLabel(i).Caption, pLabel(i), default_font)
    Next i
    
    For i = 0 To 3
        resultOption(i).Caption = get_thing("prefs", "resultoption" + Format$(i), language$, resultOption(i).Caption, resultOption(i), default_font)
    Next i
    
    pressHoldUnits(0).Caption = PU$ + "/sec"
    set_fontstuff pressHoldUnits(0), default_font
    pressHoldUnits(1).Caption = PU$ + "/min"
    set_fontstuff pressHoldUnits(1), default_font
    Check1.Caption = get_thing("prefs", "check1", language$, Check1.Caption, Check1, default_font)
    failDPDT(0).Caption = get_thing("prefs", "faildpdt0", language$, failDPDT(0).Caption, failDPDT(0), default_font)
    failDPDT(1).Caption = get_thing("prefs", "faildpdt1", language$, failDPDT(1).Caption, failDPDT(1), default_font)
    regress(0).Caption = get_thing("prefs", "regress0", language$, regress(0).Caption, regress(0), default_font)
    regress(1).Caption = get_thing("prefs", "regress1", language$, regress(1).Caption, regress(1), default_font)
    
    norefillcheck.Caption = get_thing("prefs", "norefillcheck", language$, norefillcheck.Caption, norefillcheck, default_font)
    delaycompressionliquidcheck.Caption = get_thing("prefs", "delaycompliq", language$, delaycompressionliquidcheck.Caption, delaycompressionliquidcheck, default_font)
    microflowregulatorcheck.Caption = get_thing("prefs", "uflowregcheck", language$, microflowregulatorcheck.Caption, microflowregulatorcheck, default_font)
    linSealCheck.Caption = get_thing("prefs", "linsealcheck", language$, linSealCheck.Caption, linSealCheck, default_font)
    mf_settle_check.Caption = get_thing("prefs", "mfsettle", language$, mf_settle_check.Caption, mf_settle_check, default_font)
    mf_settle_label1.Caption = get_thing("prefs", "mfsettle1", language$, mf_settle_label1.Caption, mf_settle_label1, default_font)
    mf_settle_label2.Caption = get_thing("prefs", "mfsettle2", language$, mf_settle_label2.Caption, mf_settle_label2, default_font)
    mf_temperature_check.Caption = get_thing("prefs", "mftemperature", language$, mf_temperature_check.Caption, mf_temperature_check, default_font)
    set_fontstuff Label10, default_font     ' filled during form_load
    
    For i = 0 To 2
        cflabel(i).Caption = get_thing("prefs", "cflabel" + Format$(i), language$, cflabel(i).Caption, cflabel(i), default_font)
    Next i
    
    For i = 0 To 5
        Label11(i).Caption = get_thing("prefs", "label11" + Format$(i), language$, Label11(i).Caption, Label11(i), default_font)
    Next i
    GP_avgTestCheck.Caption = get_thing("prefs", "gp_avgtestcheck", language$, GP_avgTestCheck.Caption, GP_avgTestCheck, default_font)
    Label15.Caption = get_thing("prefs", "label15", language$, Label15.Caption, Label15, default_font)
    Label16.Caption = get_thing("prefs", "label16", language$, Label16.Caption, Label16, default_font)
    permeabilityLoggingFileLabel.Caption = get_thing("prefs", "permeabilityLoggingFileLabel", language$, permeabilityLoggingFileLabel.Caption, permeabilityLoggingFileLabel, default_font)
    permeabilityLoggingCheckBox.Caption = get_thing("prefs", "permeabilityLoggingCheckBox", language$, permeabilityLoggingCheckBox.Caption, permeabilityLoggingCheckBox, default_font)
    permeabilityLoggingSelectButton.Caption = gpps2("prefs", "permeabilityLoggingSelectButton", language$, permeabilityLoggingSelectButton.Caption)
    
    Label14.Caption = get_thing("prefs", "label14", language$, Label14.Caption, Label14, default_font)
    zeroTempCheck.Caption = get_thing("prefs", "zerotemp", language$, zeroTempCheck.Caption, zeroTempCheck, default_font)
    BPTLCheck.Caption = get_thing("prefs", "bptlCheck", language$, BPTLCheck.Caption, BPTLCheck, default_font)
    BPTLSecondsLabel.Caption = get_thing("prefs", "bptlSecondsLabel", language$, BPTLSecondsLabel.Caption, BPTLSecondsLabel, default_font)
    BPTLMaxLabel.Caption = get_thing("prefs", "bptlMaxLabel", language$, BPTLMaxLabel.Caption, BPTLMaxLabel, default_font)
    BubblerCheck.Caption = get_thing("prefs", "bubblerCheck", language$, BubblerCheck.Caption, BubblerCheck, default_font)
    
    ' Other text
    ts$(1) = gpps2("prefs", "ts1", language$, "Enter unit name: (6 char)")
    ts$(2) = gpps2("prefs", "ts2", language$, "The unit name must be less than 7 characters.")
    ts$(3) = gpps2("prefs", "ts3", language$, "New Unit")
    ts$(4) = gpps2("prefs", "ts4", language$, "Enter Conversion Factor")
    ts$(5) = gpps2("prefs", "ts5", language$, "Equivalent of 1 PSI in this new unit.")
    ts$(6) = gpps2("prefs", "ts6", language$, "i.e. 1 PSI = 0.0689 BAR")
    ts$(7) = gpps2("prefs", "ts7", language$, "The unit factor must be positive")
    ts$(8) = gpps2("prefs", "ts8", language$, "Fail rate in")
    ts$(9) = gpps2("prefs", "ts9", language$, "Set Logging File")
    ts$(10) = gpps2("prefs", "ts10", language$, "Minimum Y value for pressure hold test must be smaller than maximum Y value.")
    ts$(11) = gpps2("prefs", "ts11", language$, "Maximum and minimum Y values for pressure hold test must be different.")
    ts$(12) = gpps2("prefs", "ts12", language$, "Percentage entered must be between 0 and 100")
    ts$(13) = gpps2("prefs", "ts13", language$, "Demo Mode")
    ts$(14) = gpps2("prefs", "ts14", language$, "Comm")
    ts$(15) = gpps2("prefs", "ts15", language$, "No units")
    ts$(16) = gpps2("prefs", "ts16", language$, "Percent (%)")
    ts$(17) = gpps2("prefs", "ts17", language$, "Pressure")
    ts$(18) = gpps2("prefs", "ts18", language$, "The Curve Fit routine (accessible from the Modify menu) will apply a curve-fitting algorithm to a test data file. Capwin can also be set to automatically curve-fit data files when a test is finished. (Look under the Tests tab.)")
    ts$(19) = gpps2("prefs", "ts19", language$, "Parameters for the fitting operation include the number of points to use for each fit, the maximum allowable percentage error, and the maximum distance between two adjacent points.")

End Sub


Private Sub lstInfoLines_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu frmPopupMenus.mnuPopupAdditionalInfo
    End If
End Sub

Public Sub AddInfoLine()
    Dim newValue As String
    
    newValue = InputBox("Please enter the new item:", "New Item")
    If Trim(newValue) <> "" Then
        lstInfoLines.AddItem newValue
        txtNumberOfInfoLines.Text = lstInfoLines.ListCount
    End If
End Sub

Public Sub InsertInfoLine()
    Dim newValue As String
    
    If lstInfoLines.ListIndex > -1 Then
        newValue = InputBox("Please enter the new item:", "New Item")
        If Trim(newValue) <> "" Then
            lstInfoLines.AddItem newValue, lstInfoLines.ListIndex
            txtNumberOfInfoLines.Text = lstInfoLines.ListCount
        End If
    Else
        MsgBox "Please select a location to insert value first.", vbInformation
    End If
End Sub

Public Sub DeleteInfoLine()
    Dim msgres As VbMsgBoxResult
    
    If lstInfoLines.ListIndex > -1 Then
        msgres = MsgBox("Are you sure you want to delete the following value: " & lstInfoLines.Text & "?", vbYesNo)
        If msgres = vbYes Then
            lstInfoLines.RemoveItem lstInfoLines.ListIndex
            txtNumberOfInfoLines.Text = lstInfoLines.ListCount
        End If
    Else
        MsgBox "Please select an item to delete first.", vbInformation
    End If
End Sub

Public Sub RenameInfoLine()
    Dim newValue As String
    Dim Index As Integer
    
    If lstInfoLines.ListIndex > -1 Then
        newValue = InputBox("Please enter the new value:", "New Info Item", lstInfoLines.Text)
        If Trim(newValue) <> "" Then
            Index = lstInfoLines.ListIndex
            lstInfoLines.RemoveItem Index
            lstInfoLines.AddItem newValue, Index
            
        End If
    End If
End Sub

Public Sub SetCurrentTab(currentTab As Integer)
    SSTab1.Tab = currentTab
End Sub


'Sam Bouabane 8/31/12
'these allow the user to set the starting pressure range for a pressure hold
'if the value in this box changes, set startP1 to that value and call the handler
'with the value 1 to signal the start values change
Private Sub startingP1_LostFocus()
    startThresholdMin = val(startingP1.Text)
    startingP1.Text = str$(startThresholdMin)
End Sub

'if the value in this box changes, set startP2 to that value and call the handler
'with the value 2 to signal the start values change
Private Sub startingP2_LostFocus()
    startThresholdMax = val(startingP2.Text)
    startingP2.Text = str$(startThresholdMax)
End Sub


