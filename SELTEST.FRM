VERSION 5.00
Begin VB.Form Select_Test 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Selection"
   ClientHeight    =   9750
   ClientLeft      =   9540
   ClientTop       =   3270
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "SELTEST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9750
   ScaleWidth      =   8175
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Air Resistivity Test"
      Height          =   375
      Left            =   240
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame hhSubMenu 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   4200
      TabIndex        =   89
      Top             =   2550
      Width           =   3615
      Begin VB.OptionButton optBtHydrohead 
         BackColor       =   &H0000FFFF&
         Caption         =   "Run as Hydrohead test"
         Height          =   465
         Left            =   0
         TabIndex        =   92
         Top             =   0
         Width           =   3615
      End
      Begin VB.OptionButton optBtBurst 
         BackColor       =   &H0000FFFF&
         Caption         =   "Run as Burst test"
         Height          =   465
         Left            =   0
         TabIndex        =   91
         Top             =   450
         Width           =   3615
      End
      Begin VB.OptionButton optBtMullen 
         BackColor       =   &H0000FFFF&
         Caption         =   "Run as Mullen test"
         Height          =   465
         Left            =   0
         TabIndex        =   90
         Top             =   900
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.Frame framePassFailOption 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   300
         TabIndex        =   63
         Top             =   7050
         Width           =   5895
         Begin VB.OptionButton optPassFail 
            BackColor       =   &H0000C000&
            Caption         =   "Pass/Fail According to Median Pore Size"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   65
            Top             =   480
            Width           =   3975
         End
         Begin VB.OptionButton optPassFail 
            BackColor       =   &H0000C000&
            Caption         =   "Pass/Fail According to Bubble Point"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   64
            Top             =   120
            Width           =   3975
         End
      End
      Begin VB.Frame GurleyOptFrame 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   150
         TabIndex        =   55
         Top             =   6750
         Visible         =   0   'False
         Width           =   5055
         Begin VB.OptionButton optTappi 
            BackColor       =   &H0000FFFF&
            Caption         =   "High Pressure Gurley"
            Height          =   315
            Index           =   1
            Left            =   300
            TabIndex        =   57
            Top             =   525
            Width           =   2715
         End
         Begin VB.OptionButton optTappi 
            BackColor       =   &H0000FFFF&
            Caption         =   "Sheffield Method"
            Height          =   315
            Index           =   2
            Left            =   300
            TabIndex        =   58
            Top             =   900
            Width           =   2715
         End
         Begin VB.OptionButton optTappi 
            BackColor       =   &H0000FFFF&
            Caption         =   "Standard Gurley Method"
            Height          =   315
            Index           =   0
            Left            =   300
            TabIndex        =   56
            Top             =   150
            Value           =   -1  'True
            Width           =   2715
         End
      End
      Begin VB.Frame athena_2 
         Caption         =   "Temperature 2"
         Height          =   615
         Left            =   6240
         TabIndex        =   98
         Top             =   7695
         Visible         =   0   'False
         Width           =   1575
         Begin VB.TextBox txtSetAthena2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   100
            Text            =   "0"
            Top             =   230
            Width           =   1095
         End
         Begin VB.CheckBox cbAthena2InTest 
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame athena_1 
         Caption         =   "Temperature 1"
         Height          =   615
         Left            =   6240
         TabIndex        =   95
         Top             =   7080
         Visible         =   0   'False
         Width           =   1575
         Begin VB.TextBox txtSetAthena1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   97
            Text            =   "0"
            Top             =   230
            Width           =   1095
         End
         Begin VB.CheckBox cbAthena1InTest 
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.CheckBox cbPretestFlip2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Open Valve 9"
         Height          =   375
         Left            =   2865
         TabIndex        =   94
         Top             =   8520
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.CheckBox cbPretestFlip1 
         BackColor       =   &H0000C0C0&
         Caption         =   "Open Valve 4"
         Height          =   375
         Left            =   465
         TabIndex        =   93
         Top             =   8535
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.CheckBox chk3StageTest 
         BackColor       =   &H0000C000&
         Caption         =   "Run as 3-Stage test"
         Height          =   255
         Left            =   4200
         TabIndex        =   82
         Top             =   4200
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox chkRunFrazierAsPassFail 
         BackColor       =   &H0000FF00&
         Caption         =   "Run Test As Pass/Fail"
         Height          =   315
         Left            =   750
         TabIndex        =   81
         Top             =   6150
         Width           =   6165
      End
      Begin VB.CheckBox singlePointCheck 
         BackColor       =   &H0000FF00&
         Caption         =   "Single point test"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   6720
         Width           =   6135
      End
      Begin VB.OptionButton gurley 
         BackColor       =   &H0000FFFF&
         Caption         =   "Gurley Test"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3360
         Width           =   3375
      End
      Begin VB.CheckBox ckbMultiStep 
         BackColor       =   &H0000FF00&
         Caption         =   "Use Multi-Step Test"
         Height          =   315
         Left            =   750
         TabIndex        =   49
         Top             =   5250
         Visible         =   0   'False
         Width           =   6165
      End
      Begin VB.Frame gastest 
         BackColor       =   &H0000C000&
         Caption         =   "Gas Permeability Options"
         Height          =   855
         Left            =   4080
         TabIndex        =   30
         Top             =   1560
         Width           =   3615
         Begin VB.OptionButton presdrop 
            BackColor       =   &H0000FF00&
            Caption         =   "Pressure Drop Test"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   240
            Width           =   3255
         End
         Begin VB.OptionButton normgp 
            BackColor       =   &H0000FF00&
            Caption         =   "Normal Test"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.CommandButton singlePointSetup 
         Caption         =   "Setup"
         Height          =   255
         Left            =   5880
         TabIndex        =   23
         Top             =   6720
         Width           =   1215
      End
      Begin VB.OptionButton cfp 
         BackColor       =   &H0000FFFF&
         Caption         =   "Capillary Flow Porometry"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   3375
      End
      Begin VB.OptionButton bp 
         BackColor       =   &H0000FFFF&
         Caption         =   "Bubble Point"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   3375
      End
      Begin VB.OptionButton integ 
         BackColor       =   &H0000FFFF&
         Caption         =   "Integrity Test"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   3375
      End
      Begin VB.OptionButton gasp 
         BackColor       =   &H0000FFFF&
         Caption         =   "Gas Permeametry"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3375
      End
      Begin VB.OptionButton phold 
         BackColor       =   &H0000FFFF&
         Caption         =   "Pressure Hold"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1920
         Width           =   3375
      End
      Begin VB.OptionButton microflow 
         BackColor       =   &H0000FFFF&
         Caption         =   "Microflow P vs. Time"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.OptionButton liqperm 
         BackColor       =   &H0000FFFF&
         Caption         =   "Liquid Permeametry"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2640
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.OptionButton hh 
         BackColor       =   &H0000FFFF&
         Caption         =   "Hydrohead"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton ok 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   2280
         TabIndex        =   37
         Top             =   9000
         Width           =   1215
      End
      Begin VB.CommandButton cancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4320
         TabIndex        =   38
         Top             =   9000
         Width           =   1215
      End
      Begin VB.OptionButton SurfaceArea 
         BackColor       =   &H0000FFFF&
         Caption         =   "Envelope Surface Area"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CheckBox RunPassFail 
         BackColor       =   &H0000FF00&
         Caption         =   "Run Test As Pass/Fail"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   5280
         Width           =   6135
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   1200
         TabIndex        =   31
         Top             =   4800
         Width           =   5895
         Begin VB.CommandButton Command2 
            Caption         =   "Pressure Step List Editor"
            Height          =   495
            Left            =   2160
            TabIndex        =   59
            Top             =   0
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Load Pressure Step List"
            Height          =   495
            Left            =   4080
            TabIndex        =   19
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label PListFilename 
            BackColor       =   &H0000C000&
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   5175
         End
      End
      Begin VB.CheckBox microcfp 
         BackColor       =   &H0000C000&
         Caption         =   "Microflow Porometry"
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   3840
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   750
         TabIndex        =   35
         Top             =   5550
         Width           =   5895
         Begin VB.TextBox minPassValue 
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox maxPassValue 
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox StopOnFail 
            BackColor       =   &H0000C000&
            Caption         =   "Stop if test fails (acts as an ""abort"" command)"
            Height          =   375
            Left            =   1080
            TabIndex        =   25
            Top             =   960
            Width           =   4335
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C000&
            Caption         =   "Minimum pass bubble point pore size (microns)"
            Height          =   255
            Left            =   -300
            TabIndex        =   40
            Top             =   0
            Width           =   4695
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C000&
            Caption         =   "Maximum pass bubble point pore size (microns)"
            Height          =   255
            Left            =   1080
            TabIndex        =   39
            Top             =   720
            Width           =   4575
         End
      End
      Begin VB.Frame frmFrazierPassFail 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   150
         TabIndex        =   76
         Top             =   5700
         Width           =   5895
         Begin VB.TextBox txtMinFrazierPass 
            Height          =   375
            Left            =   120
            TabIndex        =   78
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtMaxFrazierPass 
            Height          =   375
            Left            =   120
            TabIndex        =   77
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label8 
            BackColor       =   &H0000C000&
            Caption         =   "Minimum pass Frazier value"
            Height          =   255
            Left            =   1080
            TabIndex        =   80
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label Label7 
            BackColor       =   &H0000C000&
            Caption         =   "Maximum pass Frazier value"
            Height          =   255
            Left            =   1080
            TabIndex        =   79
            Top             =   720
            Width           =   4575
         End
      End
      Begin VB.Frame framePassFailParameters 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   300
         TabIndex        =   66
         Top             =   7200
         Width           =   5895
         Begin VB.TextBox txtMaxMedianPass 
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtMinMedianPass 
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label5 
            BackColor       =   &H0000C000&
            Caption         =   "Maximum pass Median pore size (microns)"
            Height          =   255
            Left            =   1080
            TabIndex        =   70
            Top             =   720
            Width           =   4575
         End
         Begin VB.Label Label4 
            BackColor       =   &H0000C000&
            Caption         =   "Minimum pass Median pore size (microns)"
            Height          =   255
            Left            =   1080
            TabIndex        =   69
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.OptionButton frazier 
         BackColor       =   &H0000FFFF&
         Caption         =   "Frazier Test"
         Height          =   375
         Left            =   120
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   3720
         Width           =   3375
      End
      Begin VB.OptionButton resin 
         BackColor       =   &H0000FFFF&
         Caption         =   "Resin Test"
         Height          =   375
         Left            =   120
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   4080
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Frame bpsub 
         BackColor       =   &H0000C000&
         Caption         =   "Bubble Point Options"
         Height          =   2415
         Left            =   4080
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
         Begin VB.Frame Frame4 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   615
            Left            =   360
            TabIndex        =   83
            Top             =   720
            Width           =   3015
            Begin VB.OptionButton optBPAutoDetectMethod 
               BackColor       =   &H0080FF80&
               Caption         =   "Detect Via F/PT"
               Height          =   315
               Index           =   0
               Left            =   150
               TabIndex        =   85
               Top             =   0
               Width           =   2565
            End
            Begin VB.OptionButton optBPAutoDetectMethod 
               BackColor       =   &H0080FF80&
               Caption         =   "Detect Via Delta F/PT"
               Height          =   315
               Index           =   1
               Left            =   150
               TabIndex        =   84
               Top             =   300
               Width           =   2565
            End
         End
         Begin VB.TextBox txtBPTestCount 
            Height          =   375
            Left            =   480
            TabIndex        =   61
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox chkBPRunMultipleTests 
            BackColor       =   &H0000FF00&
            Caption         =   "Run Multiple Tests"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1440
            Width           =   3255
         End
         Begin VB.OptionButton manual 
            BackColor       =   &H0000FF00&
            Caption         =   "Manual Detection"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   3255
         End
         Begin VB.OptionButton standard 
            BackColor       =   &H0000FF00&
            Caption         =   "Automatic Detection"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C000&
            Caption         =   "# of Tests"
            Height          =   375
            Left            =   1440
            TabIndex        =   62
            Top             =   1920
            Width           =   1695
         End
      End
      Begin VB.Frame microsub 
         BackColor       =   &H0000C000&
         Caption         =   "Microflow Options"
         Height          =   615
         Left            =   4080
         TabIndex        =   33
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CheckBox linSealCheck 
            BackColor       =   &H0000FF00&
            Caption         =   "Linear seal"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.CheckBox PListCheck 
         BackColor       =   &H0000FF00&
         Caption         =   "Use pressure step list"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   4560
         Width           =   6135
      End
      Begin VB.CommandButton btnMultiStep 
         Caption         =   "Setup"
         Height          =   255
         Left            =   5400
         TabIndex        =   50
         Top             =   5280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame cfpsub 
         BackColor       =   &H0000C000&
         Caption         =   "Capillary Flow Porometry Options"
         Height          =   3735
         Left            =   4080
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Frame lindryframe 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1095
            Left            =   360
            TabIndex        =   44
            Top             =   1680
            Visible         =   0   'False
            Width           =   3015
            Begin VB.OptionButton lindrytype 
               BackColor       =   &H0080FF80&
               Caption         =   "Darcy"
               Height          =   315
               Index           =   1
               Left            =   0
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   360
               Width           =   2895
            End
            Begin VB.OptionButton lindrytype 
               BackColor       =   &H0080FF80&
               Caption         =   "Square Root"
               Height          =   315
               Index           =   2
               Left            =   0
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   720
               Width           =   3015
            End
            Begin VB.OptionButton lindrytype 
               BackColor       =   &H0080FF80&
               Caption         =   "Linear"
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   2775
            End
         End
         Begin VB.CheckBox tobp 
            BackColor       =   &H0000C000&
            Caption         =   "Stop Dry at Bubble Point"
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   1680
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.CheckBox nowait 
            BackColor       =   &H0000C000&
            Caption         =   "No Wait at Dry"
            Height          =   375
            Left            =   360
            TabIndex        =   43
            Top             =   1680
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.CheckBox fsense 
            BackColor       =   &H0000C000&
            Caption         =   "Use Fluid Sensor"
            Height          =   375
            Left            =   360
            TabIndex        =   42
            Top             =   2040
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.OptionButton wudd 
            BackColor       =   &H0000FF00&
            Caption         =   "Wet Up/Dry Down"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   3375
         End
         Begin VB.OptionButton wudu 
            BackColor       =   &H0000FF00&
            Caption         =   "Wet Up/Dry Up"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   600
            Width           =   3375
         End
         Begin VB.OptionButton duwu 
            BackColor       =   &H0000FF00&
            Caption         =   "Dry Up/Wet Up"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   960
            Width           =   3375
         End
         Begin VB.OptionButton wuld 
            BackColor       =   &H0000FF00&
            Caption         =   "Wet Up/Calc. Dry"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1320
            Width           =   3375
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   360
            TabIndex        =   86
            Top             =   2880
            Width           =   3015
            Begin VB.OptionButton planeType 
               BackColor       =   &H0080FF80&
               Caption         =   "In-Plane"
               Height          =   315
               Index           =   1
               Left            =   0
               TabIndex        =   87
               TabStop         =   0   'False
               Top             =   360
               Width           =   3015
            End
            Begin VB.OptionButton planeType 
               BackColor       =   &H0080FF80&
               Caption         =   "Thru-Plane"
               Height          =   465
               Index           =   0
               Left            =   0
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   0
               Width           =   3015
            End
         End
      End
      Begin VB.Frame lpsub 
         BackColor       =   &H0000C000&
         Caption         =   "Liquid Permeametry Options"
         Height          =   2535
         Left            =   4080
         TabIndex        =   45
         Top             =   120
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CheckBox chkLPRunMultipleTests 
            BackColor       =   &H0000FF00&
            Caption         =   "Run Multiple Tests"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox txtLPTestCount 
            Height          =   375
            Left            =   480
            TabIndex        =   72
            Top             =   2040
            Width           =   855
         End
         Begin VB.OptionButton BuildCakeByPressure 
            BackColor       =   &H0000FF00&
            Caption         =   "Build Cake By Pressure"
            Height          =   315
            Left            =   150
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   676
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VB.OptionButton Amb 
            BackColor       =   &H0000FF00&
            Caption         =   "Ambient Pressure Test"
            Height          =   315
            Left            =   150
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   413
            Width           =   3465
         End
         Begin VB.OptionButton SlurryTubeWashCycle 
            BackColor       =   &H0000FF00&
            Caption         =   "Slurry Tube Wash Cycle"
            Height          =   315
            Left            =   150
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VB.OptionButton BuildCakeByFlow 
            BackColor       =   &H0000FF00&
            Caption         =   "Build Cake By Flow"
            Height          =   315
            Left            =   150
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   939
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VB.OptionButton elev 
            BackColor       =   &H0000FF00&
            Caption         =   "Elevated Pressure Test"
            Height          =   315
            Left            =   150
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   150
            Width           =   3465
         End
         Begin VB.Label Label6 
            BackColor       =   &H0000C000&
            Caption         =   "# of Tests"
            Height          =   255
            Left            =   1440
            TabIndex        =   74
            Top             =   2160
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "Select_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctmode%, btmode%, ltmode%
Dim ts$(4)                  ' Text strings for this form

Private Sub Amb_Click()
' User clicked "ambient" liquid perm option
    
    ' **********
    ' Begin code inserted by Tim Richards for Balance not Penetrometer on Wednesday 6/9/04
    ' ---> You can only have elevated permeability with the balance not penet setup
    If g_bBalanceNotPenet = True Then
        Amb.value = False               'no ambient
        elev.value = True               'always elev
    End If
    '
    ' End code inserted by Tim Richards 6/9/04
    ' **********
    PListCheck.value = False      '6.71.123.13
    togglePListVisibility (False)

End Sub
Private Sub cbPretestFlip1_Click()
    If cbPretestFlip1.value = 1 Then
        Move_Valve Flip1Valve, "O"
    Else
        Move_Valve Flip1Valve, "C"
    End If
End Sub

Private Sub cbPretestFlip2_Click()
    If cbPretestFlip2.value = 1 Then
        Move_Valve Flip2Valve, "O"
    Else
        Move_Valve Flip2Valve, "C"
    End If
End Sub

Private Sub bp_Click()
    
    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = True
    'edc 09-18-07
   
    microsub.Visible = False
    gastest.Visible = False
    'toggleGasMultiSet (False)
    togglePListVisibility (False)
    togglePassFailVisibility (True)
    toggleSinglePointVisibility (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    
    If RunPassFail.value = 1 Then
        passFailType(current_unit%) = 0
        Frame1.Visible = True
        framePassFailOption.Visible = False
        framePassFailParameters.Visible = False
    End If
    
    GurleyOptFrame.Visible = False
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
'    chkMullenTest.Visible = False
    
End Sub

Private Sub btnMultiStep_Click()
    prefsForm.SSTab1.Tab = 7
    prefsForm.Show 1
End Sub

Private Sub BuildCakeByFlow_Click() '6.71.123.12
    PListCheck.value = False
    togglePListVisibility (False)
End Sub

Private Sub BuildCakeByPressure_Click() '6.71.123.12
    PListCheck.value = False
    togglePListVisibility (False)
End Sub

Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub cfp_Click()
    
    bpsub.Visible = False
    'edc 09-18-07
    'toggleGasMultiSet (False)
    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = True
    microsub.Visible = False
    gastest.Visible = False
    togglePassFailVisibility (True)
    togglePListVisibility (True)
    toggleSinglePointVisibility (False)
    If DiffPG Then
        microcfp.Visible = True
        chk3StageTest.Visible = True
    End If
    
    If RunPassFail.value = 1 Then
        Frame1.Visible = False
        framePassFailOption.Visible = True
        framePassFailParameters.Visible = True
    End If
    
    
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
    GurleyOptFrame.Visible = False
    
'    chkMullenTest.Visible = False

End Sub

Private Sub chk3StageTest_Click()
    If chk3StageTest.value = 1 Then
        If microcfp.value = 1 Then
            microcfp.value = 0
        End If
    End If
End Sub

Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub microcfp_Click()
    If microcfp.value = 1 Then
        If chk3StageTest.value = 1 Then
            chk3StageTest.value = 0
        End If
    End If
End Sub

Private Sub ckbMultiStep_Click()
   'GP_multisetTest = (ckbMultiStep.value = vbChecked)
     'btnMultiStep.Visible = GP_multisetTest
End Sub

Private Sub Command1_Click()
' Change the pressure list pathname

    ' Configure and open file selection box
    fsel_path$ = EXE_Path$ + "\parms\*.psl"
    fsel_title$ = ts$(4)        ' "Choose pressure list"
    fsel_name$ = ""
    fsel_io = True
    fsel Me.hwnd
   
    If fsel_return = "" Then Exit Sub
    
    PS_path$ = fsel_return
    PListFilename.Caption = PS_path$
    
End Sub

Private Sub Command2_Click()
    Me.Enabled = False
    stepListForm.Show 1
    Me.Enabled = True
    
End Sub

Private Sub duwu_Click()

    tobp.Visible = False
    If fluidsensor Then
        fsense.Visible = True
    End If
    nowait.Visible = False
    lindryframe.Visible = False
    
End Sub

Private Sub elev_Click()
    togglePListVisibility (True)
End Sub

Private Sub Form_Load()

LoadTextStrings

Me.Left = (Screen.Width - Me.Width) / 2
Me.top = (Screen.Height - Me.Height) / 2
Me.gastest.Visible = False
'GP_singlePointTest = False

'
'If AirResistivity Then
'    Option1.Visible = True
'End If

' Show pretest valve flips - Mix
If Flip1Show Then
    cbPretestFlip1.Caption = Flip1Name
    cbPretestFlip1.Visible = True
    Move_Valve Flip1Valve, "C"
End If
If Flip2Show Then
    cbPretestFlip2.Caption = Flip2Name
    cbPretestFlip2.Visible = True
    Move_Valve Flip2Valve, "C"
End If

' Show Athena controls
If UseAthena1 Then
    athena_1.Visible = True
    cbAthena1InTest.value = IIf(Athena1InTest, 1, 0)
    txtSetAthena1.Text = CStr(Athena1Target)
End If
If UseAthena2 Then
    athena_2.Visible = True
    cbAthena2InTest.value = IIf(Athena2InTest, 1, 0)
    txtSetAthena2.Text = CStr(Athena2Target)
End If

If GasPerm Then
    'edc 09-18-07
    'toggleGasMultiSet (True)
    cfp.Visible = False
    bp.Visible = False
    integ.Visible = False
    togglePassFailVisibility (False)
    toggleSinglePointVisibility (True)
    
End If

If liqpermonly Then
    cfp.Visible = False
    bp.Visible = False
        'edc 09-18-07
    'toggleGasMultiSet (True)
    integ.Visible = False
    gasp.Visible = False
    frazier.Visible = False
    SurfaceArea.Visible = False
    phold.Visible = False
    togglePassFailVisibility (False)
    '6.71.123.12 If ltmode% = 1 Then
    '6.71.123.12    togglePListVisibility (False)
    '6.71.123.12    ' Only visible for elevated
    '6.71.123.12   togglePListVisibility (True)
    '6.71.123.12 End If
    
    'PList Only visible for elevated
    togglePListVisibility (ltmode% = 2) '6.71.123.12

    toggleSinglePointVisibility (False)
    
    gurley.Visible = False
    frazier.Visible = False
End If

If itester Then
    cfp.Visible = False
    gasp.Visible = False
    frazier.Visible = False
    gurley.Visible = False
    
        'edc 09-18-07
    'toggleGasMultiSet (False)
    SurfaceArea.Visible = False
    Rem check for types of tests that can't be done
    If TType%(current_unit%) = 1 Or TType%(current_unit%) = 3 Then
        TType%(current_unit%) = 2
        TMode%(current_unit%) = 1
    End If
    toggleSinglePointVisibility (False)
End If

If BPTester Then
    cfp.Visible = False
    gasp.Visible = False
    frazier.Visible = False
    gurley.Visible = False
    
        'edc 09-18-07
    'toggleGasMultiSet (False)
    SurfaceArea.Visible = False
    If TType%(current_unit%) = 5 Then
        ' allow pressure hold test in bubble point testers
        ' just make sure that the tmode is sane
        ' This is probably not necessary, since TMode isn't really used for pressure hold
        TMode%(current_unit%) = 1
    ElseIf TType%(current_unit%) <> 2 Then
        ' if the last test was not pressure hold or bubble point...
        ' set it to the standard bubble point test
        TType%(current_unit%) = 2
        TMode%(current_unit%) = 1
    ElseIf TMode%(current_unit%) > 1 Then
        ' 6.71.38x begin
        ' if they have a bubblepoint-type test selected but it is
        ' not mode 0 or 1 (or 3 if hydrohead is enabled) set it back to 1
        If TMode%(current_unit%) <> 3 Or hydrohead = False Then
            TMode%(current_unit%) = 1
        End If
        ' 6.71.38x end
    End If
    'phold.Visible = False
    togglePListVisibility (False)
    'toggleGasMultiSet (False)
    togglePassFailVisibility (True)
    toggleSinglePointVisibility (False)
    gurley.Visible = False
    frazier.Visible = False
End If

If Not integrity Then integ.Visible = False

If DiffPG Then
    microflow.Visible = True
End If

If lvperm_enable Then
    microflow.Caption = ts$(1)          ' "liq. vapor perm."
    microflow.Visible = True
End If

If H2OPERM Then
    liqperm.Visible = True
    BuildCakeByPressure.Visible = slurry_tube_exists '6.71.123.12
    BuildCakeByFlow.Visible = slurry_tube_exists     '6.71.123.12
    SlurryTubeWashCycle.Visible = slurry_tube_exists '6.71.123.12
End If

If hydrohead Then
    hh.Visible = True
    ' 6.71.38x begin
    If hydrohead_exclusive Then
        ' BPTester variable will have already turned off most of the other stuff
        ' we only need to turn off the bubble point test selection
        bp.Visible = False
        ' only type of test is hydrohead
        TType%(current_unit%) = 2
        TMode%(current_unit%) = 3
    End If
    
    ' 6.71.38x end
End If

If Resin_Diverter_Valve > 0 Then
    resin.Visible = True
Else
    resin.Visible = False
End If

If IS_ACPP Then
    cfp.Visible = False
    bp.Visible = False
    integ.Visible = False
    gasp.Visible = False
    SurfaceArea.Visible = False
    microflow.Visible = False
    liqperm.Visible = False
    hh.Visible = False
    gurley.Visible = False
    frazier.Visible = False
    resin.Visible = False
End If
If Not WESA_enabled Then SurfaceArea.Visible = False

ctmode% = 1
btmode% = 1
ltmode% = 1

Select Case TType%(current_unit%)
    Case 1
        ctmode% = TMode%(current_unit%)
        cfp.value = True
        '6.71.68: The next two lines are redundant, because setting cfp.Value calls the cfp_Click routine,
        ' and cfp_Click makes the same assignments.
        'togglePListVisibility (True)
        'toggleSinglePointVisibility (False)
        
    Case 2
        ' if they call for integrity test and we can't do it
        ' switch to bubble point test, auto detection mode
        If TMode%(current_unit%) = 2 And Not integrity Then TMode%(current_unit%) = 1
        Select Case TMode%(current_unit%)
        Case 0, 1
            btmode% = TMode%(current_unit%)
            bp.value = True
            togglePListVisibility (False)
            togglePassFailVisibility (True)
            toggleSinglePointVisibility (False)
            chkBPRunMultipleTests.value = IIf(bpRunMultipleTests, 1, 0)
            txtBPTestCount.Text = bpTestCount
        Case 2
            integ.value = True
        Case 3
            hh.value = True
                If (hhRunAsHydrohead(current_unit%)) Then
                    optBtHydrohead.value = True
                ElseIf (hhRunAsMullen(current_unit%)) Then
                    optBtMullen.value = True
                Else
                    optBtBurst.value = True
                End If
                
        End Select
        'toggleGasMultiSet (False)
    Case 3
        Select Case TMode%(current_unit%)
        Case 1
            SurfaceArea.value = True
        Case 10
            gurley.value = True
            optTappi(0).value = True
        Case 11
            gurley.value = True
            optTappi(1).value = True
        Case 12
            gurley.value = True
            optTappi(2).value = True
        Case 20
            frazier.value = True
            If runFrazierAsPassFail(current_unit%) Then
                chkRunFrazierAsPassFail.value = vbChecked
                frmFrazierPassFail.Visible = True
            Else
                frmFrazierPassFail.Visible = False
            End If
        Case Else
            gasp.value = True
            If PDrop Then Me.gastest.Visible = True
            If (TMode%(current_unit%) = 2) And PDrop Then
                Me.presdrop.value = True
            Else
                Me.normgp.value = True
                TMode%(current_unit%) = 0
            End If
            togglePListVisibility (True)
            togglePassFailVisibility (False)
            toggleSinglePointVisibility (True)
        End Select
        'toggleGasMultiSet (True)
    Case 4
        ltmode% = TMode%(current_unit%)
        liqperm.value = True
        togglePassFailVisibility (False)
        togglePListVisibility (ltmode% = 2)
        toggleSinglePointVisibility (False)
        'toggleGasMultiSet (False)
    Case 5
        phold.value = True
        togglePassFailVisibility (False)
        togglePListVisibility (False)
        'toggleGasMultiSet (False)
        toggleSinglePointVisibility (False)
    Case 6
        microflow.value = True
        togglePassFailVisibility (False)
        togglePListVisibility (False)
        toggleSinglePointVisibility (False)
        'toggleGasMultiSet (False)
    Case 7
        resin.value = True
End Select

If diffpgflow(current_unit%) Then
    microcfp.value = 1
End If

If threestagetest(current_unit%) Then
    chk3StageTest.value = 1
End If

Select Case ctmode%
    Case 1, 5
        wudd.value = True
    Case 2, 6, 8, 9
        wudu.value = True
    Case 3, 7
        duwu.value = True
    Case 4, 10, 11
        wuld.value = True
End Select

Select Case ctmode%
    Case 5
        tobp.value = 1
    Case 6, 7, 9
        fsense.value = 1
End Select

If ctmode% > 7 And ctmode% < 10 Then
    nowait.value = 1
End If

If ctmode% = 10 Then
    lindrytype(1).value = True
End If

If ctmode% = 11 Then
    lindrytype(2).value = True
End If

If btmode% = 0 Then
  manual.value = True
Else
  standard.value = True
End If

' 6.71.68: Added the "if TType%(current_unit%) ..." clause. Setting the Value variable
' performs the Click action, which sets some things we don't want set unless we're actually
' looking at a liquid perm test.
If TType%(current_unit%) = 4 Then
    If ltmode% = 1 Then
        Amb.value = True
    '6.71.123.12 Else
    ElseIf ltmode% = 2 Then                 '6.71.123.12
        elev.value = True
    ElseIf ltmode% = 13 Then
        BuildCakeByPressure.value = True
    ElseIf ltmode% = 14 Then
        BuildCakeByFlow.value = True
    ElseIf ltmode% = 15 Then
        SlurryTubeWashCycle.value = True
    Else
        MsgBox ("Invalid liquid perm choice")
    End If
End If

If unitnumber <> 0 Then
    Me.Caption = Me.Caption + " - " + ts$(2) + str$(unitnumber)       ' "Unit"
End If

optPassFail(passFailType(current_unit%)).value = True
minPassValue.Text = str$(minPassDiameter(current_unit%))
maxPassValue.Text = str$(maxPassDiameter(current_unit%))
txtMinMedianPass.Text = str$(minMedianPass(current_unit%))
txtMaxMedianPass.Text = str$(maxMedianPass(current_unit%))

optBPAutoDetectMethod(BP_AutoDetectMethod).value = True

planeType(CF_SampleType).value = True

If runAsPassFail(current_unit%) Then
    RunPassFail.value = vbChecked
    If passFailType(current_unit%) = 0 Then
        Frame1.Visible = True
        framePassFailParameters.Visible = False
    Else
        Frame1.Visible = False
        framePassFailParameters.Visible = True
    End If
    framePassFailOption.Visible = True
Else
    framePassFailOption.Visible = False
    Frame1.Visible = False
    framePassFailParameters.Visible = False
End If

If hhRunAsHydrohead(current_unit%) Then
    optBtHydrohead.value = True
End If

If hhRunAsMullen(current_unit%) Then
    optBtMullen.value = True
End If

If hhRunAsBurst(current_unit%) Then
    optBtBurst.value = True
End If


txtMinFrazierPass.Text = str$(minFrazierPass(current_unit%))
txtMaxFrazierPass.Text = str$(maxFrazierPass(current_unit%))

If stopTestOnFail(current_unit%) Then
    StopOnFail.value = vbChecked
End If

linSealCheck.value = IIf(MF_linearSeal, vbChecked, vbUnchecked)
PListCheck.value = IIf(PS_usingList, vbChecked, vbUnchecked)
'singlePointCheck.value = IIf(GP_singlePointTest, vbChecked, vbUnchecked)
If TMode%(current_unit%) > 9 And TMode%(current_unit%) < 13 Then
    singlePointCheck.value = False
Else
    singlePointCheck.value = IIf(GP_singlePointTest, vbChecked, vbUnchecked)
End If
'ckbMultiStep.value = IIf(GP_multisetTest, vbChecked, vbUnchecked)
'edc 12-11-06 alter border color and caption
Me.Caption = Me.Caption & "    " & SubCaption
Me.BackColor = lngBorderColor
If IS_ACPP Then
    cfp.Visible = False
    bp.Visible = False
    integ.Visible = False
    gasp.Visible = False
    SurfaceArea.Visible = False
    microflow.Visible = False
    liqperm.Visible = False
    hh.Visible = False
    gurley.Visible = False
    frazier.Visible = False
    resin.Visible = False
End If
End Sub

Private Sub frazier_Click()

    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = False
    gastest.Visible = False
    microsub.Visible = False
    'toggleGasMultiSet (False)
    togglePassFailVisibility (False)
    togglePListVisibility (False)
    toggleSinglePointVisibility (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    chkRunFrazierAsPassFail.Visible = True
    frmFrazierPassFail.Visible = IIf(chkRunFrazierAsPassFail.value = vbChecked, True, False)
    GurleyOptFrame.Visible = False

'    chkMullenTest.Visible = False
'    chkBurstTest.Visible = False
End Sub

Private Sub gasp_Click()

    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = False
    microsub.Visible = False
    GurleyOptFrame.Visible = False
    If PDrop Then
        Me.gastest.Visible = True
        Me.normgp.value = True
    End If
    'toggleGasMultiSet (True)
    togglePassFailVisibility (False)
    togglePListVisibility (True)
    toggleSinglePointVisibility (True)
    If DiffPG Then
        microcfp.Visible = True
        chk3StageTest.Visible = True
    End If
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
'    chkMullenTest.Visible = False
'    chkBurstTest.Visible = False

End Sub

Private Sub gurley_Click()

    GurleyOptFrame.Visible = True
    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = False
    gastest.Visible = False
    microsub.Visible = False
    'toggleGasMultiSet (False)
    togglePassFailVisibility (False)
    togglePListVisibility (False)
    toggleSinglePointVisibility (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
'    chkMullenTest.Visible = False
'    chkBurstTest.Visible = False
    
End Sub

Private Sub hh_Click()

    hhSubMenu.Visible = True
    lpsub.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = False
    gastest.Visible = False
    microsub.Visible = False
    'toggleGasMultiSet (False)
    togglePassFailVisibility (False)
    togglePListVisibility (False)
    toggleSinglePointVisibility (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
    GurleyOptFrame.Visible = False
    optBtHydrohead.Visible = True

End Sub

Private Sub integ_Click()
    
    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = False
    microsub.Visible = False
    gastest.Visible = False
    'toggleGasMultiSet (False)
    togglePassFailVisibility (False)
    togglePListVisibility (False)
    toggleSinglePointVisibility (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
    GurleyOptFrame.Visible = False
    
'    chkMullenTest.Visible = False
'    chkBurstTest.Visible = False

End Sub

Private Sub linSealCheck_Click()

    MF_linearSeal = (linSealCheck.value = vbChecked)
    WPPS Curr_U$, "mf_linear_seal", IIf(MF_linearSeal, "Y", "N"), IFile$

End Sub

Private Sub liqperm_Click()

    cfpsub.Visible = False
    hhSubMenu.Visible = False
    bpsub.Visible = False
    lpsub.Visible = True
    gastest.Visible = False
    microsub.Visible = False
    'toggleGasMultiSet (False)
    togglePassFailVisibility (False)
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
    
    '6.71.123.12 If ltmode% = 1 Then
    '6.71.123.12      togglePListVisibility (False)
    '6.71.123.12 Else
    '6.71.123.12     togglePListVisibility (True)
    '6.71.123.12 End If
    
    'PList Only visible for elevated
    togglePListVisibility (ltmode% = 2) '6.71.123.12
    
    toggleSinglePointVisibility (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    GurleyOptFrame.Visible = False

    'toggleGasMultiSet (False)
' **********
' Begin code inserted by Tim Richards for Balance Not Penet on Weds 6/9/04
'
    If g_bBalanceNotPenet = True Then
        Amb.Visible = False
        Amb.value = False
        elev.value = True
    End If
    
'
' End code inserted by Tim Richards 6/9/04
' **********
    Amb.value = True
    'gurley.Visible = False
'    chkMullenTest.Visible = False
'    chkBurstTest.Visible = False
    
End Sub

Private Sub microflow_Click()
    
    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = False
    gastest.Visible = False
    microsub.Visible = True
    togglePassFailVisibility (False)
    togglePListVisibility (False)
    toggleSinglePointVisibility (False)
    'toggleGasMultiSet (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
    GurleyOptFrame.Visible = False

'    chkMullenTest.Visible = False
'    chkBurstTest.Visible = False

End Sub

Private Sub optPassFail_Click(Index As Integer)
    If Index = 0 Then
        Frame1.Visible = True
        framePassFailParameters.Visible = False
    Else
        framePassFailParameters.Visible = True
        Frame1.Visible = False
    End If
    passFailType(current_unit%) = Index
End Sub

Private Sub PListCheck_Click()
' Toggle visibility of pathname and set global flag
    
    Frame2.Visible = (PListCheck.value = vbChecked)
    PListFilename.Caption = PS_path$
    
End Sub

Private Sub resin_Click()
    GurleyOptFrame.Visible = False
    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = False
    gastest.Visible = False
    microsub.Visible = False
    'toggleGasMultiSet (False)
    togglePassFailVisibility (False)
    togglePListVisibility (False)
    toggleSinglePointVisibility (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
    GurleyOptFrame.Visible = False

'    chkMullenTest.Visible = False
'    chkBurstTest.Visible = False
    
End Sub

Private Sub RunPassFail_Click()
    
    If phold.value Then         ' Hide options non-applicable to press hold test
        Label1.Visible = False
        Label2.Visible = False
        minPassValue.Visible = False
        maxPassValue.Visible = False
    Else
        Label1.Visible = True
        Label2.Visible = True
        minPassValue.Visible = True
        maxPassValue.Visible = True
    End If
    
    If RunPassFail.value Then
        runAsPassFail(current_unit%) = True
        If bp.value = True Then
            passFailType(current_unit%) = 0
            Frame1.Visible = True
        Else
            framePassFailOption.Visible = True
            If passFailType(current_unit%) = 0 Then
                Frame1.Visible = True
            Else
                framePassFailParameters.Visible = True
            End If
        End If
    Else
        framePassFailOption.Visible = False
        Frame1.Visible = False
        framePassFailParameters.Visible = False
        runAsPassFail(current_unit%) = False
    End If
    
End Sub

Private Sub singlePointCheck_Click()
    GP_singlePointTest = (singlePointCheck.value = vbChecked)
    singlePointSetup.Visible = GP_singlePointTest
End Sub

Private Sub singlePointSetup_Click()

    prefsForm.SSTab1.Tab = 7
    prefsForm.Show 1
    
    
End Sub

Private Sub SlurryTubeWashCycle_Click() '6.71.123.12
    PListCheck.value = False
    togglePListVisibility (False)
End Sub

Private Sub StopOnFail_Click()
        
    stopTestOnFail(current_unit%) = IIf(StopOnFail.value, True, False)

End Sub

Private Sub surfacearea_click()

    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = False
    microsub.Visible = False
    gastest.Visible = False
    'toggleGasMultiSet (False)
    togglePassFailVisibility (False)
    togglePListVisibility (False)
    toggleSinglePointVisibility (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    GurleyOptFrame.Visible = False
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
 '   chkMullenTest.Visible = False
 '   chkBurstTest.Visible = False

    
End Sub

Private Sub ok_Click()
    Dim r As Long
    If Not gasp Then GP_multiAverageTest = False

    If cfp.value Then
        TType%(current_unit%) = 1
        If wudd.value Then
            If tobp.value = 1 Then
                TMode%(current_unit%) = 5
            Else
                TMode%(current_unit%) = 1
            End If
        ElseIf wuld.value Then
            If lindrytype(1).value = True Then
                TMode%(current_unit%) = 10
            ElseIf lindrytype(2).value = True Then
                TMode%(current_unit%) = 11
            Else
                TMode%(current_unit%) = 4
            End If
        ElseIf duwu.value Then
            If fsense.value = 1 Then
                TMode%(current_unit%) = 7
            Else
                TMode%(current_unit%) = 3
            End If
        Else
            Rem must be wudu
            If nowait.value = 1 Then
                If fsense.value = 1 Then
                    TMode%(current_unit%) = 9
                Else
                    TMode%(current_unit%) = 8
                End If
            Else
                If fsense.value = 1 Then
                    TMode%(current_unit%) = 6
                Else
                    TMode%(current_unit%) = 2
                End If
            End If
        End If
        If microcfp.value = 1 Then
            diffpgflow(current_unit%) = True
        Else
            diffpgflow(current_unit%) = False
        End If
        If chk3StageTest.value = 1 Then
            threestagetest(current_unit%) = True
        Else
            threestagetest(current_unit%) = False
        End If
        If planeType(0).value Then
            CF_SampleType = 0
        Else
            CF_SampleType = 1
        End If
    ElseIf bp.value Then
        TType%(current_unit%) = 2
        If manual.value Then
            TMode%(current_unit%) = 0
        Else
            TMode%(current_unit%) = 1
        End If
        If chkBPRunMultipleTests.value = 1 Then
            bpRunMultipleTests = True
            bpTestCount = val(txtBPTestCount.Text)
        Else
            bpRunMultipleTests = False
        End If
        If optBPAutoDetectMethod(0).value Then
            BP_AutoDetectMethod = 0
        Else
            BP_AutoDetectMethod = 1
        End If
    ElseIf integ.value Then
        TType%(current_unit%) = 2
        TMode%(current_unit%) = 2
    ElseIf gasp.value Then
        TType%(current_unit%) = 3
        TMode%(current_unit%) = 0
        If Me.presdrop.value Then TMode%(current_unit%) = 2
        If microcfp.value = 1 Then
            diffpgflow(current_unit%) = True
        Else
            diffpgflow(current_unit%) = False
        End If
        If chk3StageTest.value = 1 Then
            threestagetest(current_unit%) = True
        Else
            threestagetest(current_unit%) = False
        End If
        GP_singlePointTest = (singlePointCheck.value = vbChecked)
    ElseIf gurley.value Then
        'Should be same as Gas Permeability (gasp)
        TType%(current_unit%) = 3
        If optTappi(0).value Then
            TMode%(current_unit%) = 10
            GP_target = 0.177
            GP_delay = 1
            GP_duration = 10
            GP_interval = 1
            GP_numavg = 1
        ElseIf optTappi(1).value Then
            TMode%(current_unit%) = 11
            GP_target = 0.435
            GP_delay = 1
            GP_duration = 10
            GP_interval = 1
            GP_numavg = 1
        Else
            TMode%(current_unit%) = 12
            GP_target = 1.43
            GP_delay = 1
            GP_duration = 10
            GP_interval = 1
            GP_numavg = 1
        End If
        Diam(current_unit%) = (1.125 * linear_unit_conversion#)
        diffpgflow(current_unit%) = False
        threestagetest(current_unit%) = False
        GP_singlePointTest = True
    ElseIf frazier.value Then
        TType%(current_unit%) = 3
        TMode%(current_unit%) = 20
        GP_target = 0.0181
        GP_delay = 1
        GP_duration = 10
        GP_interval = 1
        GP_numavg = 1
        diffpgflow(current_unit%) = False
        threestagetest(current_unit%) = False
        GP_singlePointTest = True
        WPPS Curr_U$, "gp_target", str$(GP_target), IFile$
        WPPS Curr_U$, "gp_delay", str$(GP_delay), IFile$
        WPPS Curr_U$, "gp_duration", str$(GP_duration), IFile$
        WPPS Curr_U$, "gp_interval", str$(GP_interval), IFile$
        WPPS Curr_U$, "gp_numavg", str$(GP_numavg), IFile$
        WPPS Curr_U$, "gp_singlepoint", IIf(GP_singlePointTest, "Y", "N"), IFile$
        If chkRunFrazierAsPassFail.value = vbChecked Then
            runFrazierAsPassFail(current_unit%) = True
            minFrazierPass(current_unit%) = val(txtMinFrazierPass.Text)
            maxFrazierPass(current_unit%) = val(txtMaxFrazierPass.Text)
            WPPS Curr_U$, "runFrazierAsPassFail", "1", IFile$
            WPPS Curr_U$, "minFrazierPass", str$(minFrazierPass(current_unit%)), IFile$
            WPPS Curr_U$, "maxFrazierPass", str$(maxFrazierPass(current_unit%)), IFile$
        Else
            runFrazierAsPassFail(current_unit%) = False
            WPPS Curr_U$, "runFrazierAsPassFail", "0", IFile$
        End If
    ElseIf resin.value Then
        TType%(current_unit%) = 7
    ElseIf SurfaceArea.value Then
        TType%(current_unit%) = 3
        TMode%(current_unit%) = 1
    ElseIf phold.value Then
        TType%(current_unit%) = 5
    ElseIf microflow.value Then
        TType%(current_unit%) = 6
    ElseIf liqperm.value Then
        TType%(current_unit%) = 4
        If Amb.value Then
            TMode%(current_unit%) = 1
        ElseIf elev.value Then                 '6.71.123.12
            TMode%(current_unit%) = 2
        ElseIf BuildCakeByPressure.value Then
            TMode%(current_unit%) = 13
        ElseIf BuildCakeByFlow.value Then
            TMode%(current_unit%) = 14
        ElseIf SlurryTubeWashCycle.value Then
            TMode%(current_unit%) = 15
        Else
            If minPassDiameter(current_unit%) < 0 Or maxPassDiameter(current_unit%) < minPassDiameter(current_unit%) Or minPassDiameter(current_unit%) > maxPassDiameter(current_unit%) Then
                MsgBox ("Invalid liquid perm choice")
                MsgBox ts$(3)       ' ("Illegal pass/fail values. Please reenter.")
                Exit Sub
            End If
            '6.71.123.12     TMode%(current_unit%) = 2
        End If
    ElseIf hh.value Then
        TType%(current_unit%) = 2
        TMode%(current_unit%) = 3
        diffpgflow(current_unit%) = False
        threestagetest(current_unit%) = False
        ' Choose among Hydrohead, Burst, and Mullen
        
        If optBtHydrohead.value = True Then
            WPPS Curr_U$, "hhRunAsHydrohead", "1", IFile$
            WPPS Curr_U$, "hhRunAsMullen", "0", IFile$
            WPPS Curr_U$, "hhRunAsBurst", "0", IFile$
            hhRunAsBurst(current_unit%) = False
            hhRunAsMullen(current_unit%) = False
            hhRunAsHydrohead(current_unit%) = True
        ElseIf optBtBurst.value = True Then
            WPPS Curr_U$, "hhRunAsHydrohead", "0", IFile$
            WPPS Curr_U$, "hhRunAsMullen", "0", IFile$
            WPPS Curr_U$, "hhRunAsBurst", "1", IFile$
            hhRunAsBurst(current_unit%) = True
            hhRunAsMullen(current_unit%) = False
            hhRunAsHydrohead(current_unit%) = False
        Else
            WPPS Curr_U$, "hhRunAsMullen", "1", IFile$
            WPPS Curr_U$, "hhRunAsBurst", "0", IFile$
            WPPS Curr_U$, "hhRunAsHydrohead", "0", IFile$
            hhRunAsBurst(current_unit%) = False
            hhRunAsMullen(current_unit%) = True
            hhRunAsHydrohead(current_unit%) = False
        End If
        
    End If
    If Option1.value Then
        TType%(current_unit%) = 8
        Unload Me
    End If
        
    If runAsPassFail(current_unit%) Then
        minPassDiameter(current_unit%) = val(minPassValue.Text)
        maxPassDiameter(current_unit%) = val(maxPassValue.Text)
        minMedianPass(current_unit%) = val(txtMinMedianPass.Text)
        maxMedianPass(current_unit%) = val(txtMaxMedianPass.Text)
        If minPassDiameter(current_unit%) < 0 Or maxPassDiameter(current_unit%) < minPassDiameter(current_unit%) Or minPassDiameter(current_unit%) > maxPassDiameter(current_unit%) Then
            MsgBox ts$(3)       ' ("Illegal pass/fail values. Please reenter.")
            Exit Sub
        End If
        If minMedianPass(current_unit%) < 0 Or maxMedianPass(current_unit%) < minMedianPass(current_unit%) Or minMedianPass(current_unit%) > maxMedianPass(current_unit%) Then
            MsgBox ts$(3)       ' ("Illegal pass/fail values. Please reenter.")
            Exit Sub
        End If
    End If
    
    ' Save / activate the Athena values
    If UseAthena1 Then
        Athena1InTest = IIf(cbAthena1InTest.value = 1, True, False)
        WPPS "Capstuff", "Athena1InTest", IIf(Athena1InTest, "Y", "N"), CSFile$
        
        On Error GoTo OKClickFinalize
        If Athena1InTest Then Set_Athena Athena1Channel, val(txtSetAthena1.Text)
    End If
    If UseAthena2 Then
        Athena2InTest = IIf(cbAthena2InTest.value = 1, True, False)
        WPPS "Capstuff", "Athena2InTest", IIf(Athena2InTest, "Y", "N"), CSFile$
        
        On Error GoTo OKClickFinalize
        If Athena2InTest Then Set_Athena Athena2Channel, val(txtSetAthena2.Text)
    End If

OKClickFinalize:
    '6.71.123.12 PS_usingList = ((PListCheck.value = vbChecked) And (cfp.value Or gasp.value Or (liqperm.value And Not Amb.value)))
    '6.71.123.12 Not Amb no longer means = elev because there are three new choices.
    PS_usingList = ((PListCheck.value = vbChecked) And (cfp.value Or gasp.value Or (liqperm.value And elev.value)))
    Unload Me
End Sub

Private Sub phold_Click()

    lpsub.Visible = False
    hhSubMenu.Visible = False
    cfpsub.Visible = False
    bpsub.Visible = False
    gastest.Visible = False
    microsub.Visible = False
    'toggleGasMultiSet (False)
    togglePassFailVisibility (False)
    togglePListVisibility (False)
    toggleSinglePointVisibility (False)
    microcfp.Visible = False
    chk3StageTest.Visible = False
    GurleyOptFrame.Visible = False
    chkRunFrazierAsPassFail.Visible = False
    frmFrazierPassFail.Visible = False
'    chkMullenTest.Visible = False

End Sub

Private Sub wudd_Click()

    If fluidsensor Then
      fsense.Visible = True
    End If
    tobp.Visible = True
    nowait.Visible = False
    lindryframe.Visible = False

End Sub

Private Sub wudu_Click()

tobp.Visible = False
nowait.Visible = True
If fluidsensor Then
  fsense.Visible = True
End If
lindryframe.Visible = False

End Sub

Private Sub wuld_Click()

tobp.Visible = False
fsense.Visible = False
nowait.Visible = False
lindryframe.Visible = True

End Sub

Private Sub togglePassFailVisibility(state As Boolean)
' Set the pass/fail test options visible or invisible

    RunPassFail.Visible = state
    Frame1.Visible = state And (RunPassFail.value = vbChecked)

End Sub

Private Sub togglePListVisibility(state As Boolean)
' Set the pressure list test options visible or invisible or checked. Also set pathname.

    PListCheck.Visible = state
    Frame2.Visible = state And (PListCheck.value = vbChecked)
    
    If Frame2.Visible Then PListFilename.Caption = PS_path$
    
End Sub
'Private Sub toggleGasMultiSet(state As Boolean)
'edc 09-18-07 changes the visability of the MultiStep button and check box
'ckbMultiStep.Visible = state
'btnMultiStep.Visible = False
'End Sub

Private Sub toggleSinglePointVisibility(state As Boolean)
' Set the "single point" gas permeability options visible or invisible.
    singlePointCheck.Visible = state
    singlePointSetup.Visible = state
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    Select_Test.Caption = gpps2("selecttest", "window title", language$, Select_Test.Caption)
    cfp.Caption = get_thing("selecttest", "cfp", language$, cfp.Caption, cfp, default_font)
    bp.Caption = get_thing("selecttest", "bp", language$, bp.Caption, bp, default_font)
    integ.Caption = get_thing("selecttest", "integ", language$, integ.Caption, integ, default_font)
    gasp.Caption = get_thing("selecttest", "gasp", language$, gasp.Caption, gasp, default_font)
    SurfaceArea.Caption = get_thing("selecttest", "surfacearea", language$, SurfaceArea.Caption, SurfaceArea, default_font)
    phold.Caption = get_thing("selecttest", "phold", language$, phold.Caption, phold, default_font)
    microflow.Caption = get_thing("selecttest", "microflow", language$, microflow.Caption, microflow, default_font)
    liqperm.Caption = get_thing("selecttest", "liqperm", language$, liqperm.Caption, liqperm, default_font)
    hh.Caption = get_thing("selecttest", "hh", language$, hh.Caption, hh, default_font)
    cfpsub.Caption = get_thing("selecttest", "cfpsub", language$, cfpsub.Caption, cfpsub, default_font)
   ' ShowResult.Caption = get_thing("selecttest", "showresult", language$, ShowResult.Caption, ShowResult, default_font)
    wudd.Caption = get_thing("selecttest", "wudd", language$, wudd.Caption, wudd, default_font)
    wudu.Caption = get_thing("selecttest", "wudu", language$, wudu.Caption, wudu, default_font)
    duwu.Caption = get_thing("selecttest", "duwu", language$, duwu.Caption, duwu, default_font)
    wuld.Caption = get_thing("selecttest", "wuld", language$, wuld.Caption, wuld, default_font)
    Amb.Caption = get_thing("selecttest", "amb", language$, Amb.Caption, Amb, default_font)
    bpsub.Caption = get_thing("selecttest", "bpsub", language$, bpsub.Caption, bpsub, default_font)
    cancel.Caption = gpps2("selecttest", "cancel", language$, cancel.Caption)
    set_fontname cancel, default_font
    elev.Caption = get_thing("selecttest", "elev", language$, elev.Caption, elev, default_font)
    fsense.Caption = get_thing("selecttest", "fsense", language$, fsense.Caption, fsense, default_font)
    gastest.Caption = get_thing("selecttest", "gastest", language$, gastest.Caption, gastest, default_font)
    lindrytype(0).Caption = get_thing("selecttest", "lindry0", language$, lindrytype(0).Caption, lindrytype(0), default_font)
    lindrytype(1).Caption = get_thing("selecttest", "lindry1", language$, lindrytype(1).Caption, lindrytype(1), default_font)
    lindrytype(2).Caption = get_thing("selecttest", "lindry2", language$, lindrytype(2).Caption, lindrytype(2), default_font)
    lpsub.Caption = get_thing("selecttest", "lpsub", language$, lpsub.Caption, lpsub, default_font)
    hhSubMenu.Caption = get_thing("selecttest", "hhSubMenu", language$, hhSubMenu.Caption, hhSubMenu, default_font)
    manual.Caption = get_thing("selecttest", "manual", language$, manual.Caption, manual, default_font)
    microcfp.Caption = get_thing("selecttest", "microcfp", language$, microcfp.Caption, microcfp, default_font)
    normgp.Caption = get_thing("selecttest", "normgp", language$, normgp.Caption, normgp, default_font)
    nowait.Caption = get_thing("selecttest", "nowait", language$, nowait.Caption, nowait, default_font)
    ok.Caption = gpps2("selecttest", "ok", language$, ok.Caption)
    set_fontname ok, default_font
    presdrop.Caption = get_thing("selecttest", "presdrop", language$, presdrop.Caption, presdrop, default_font)
    standard.Caption = get_thing("selecttest", "standard", language$, standard.Caption, standard, default_font)
    tobp.Caption = get_thing("selecttest", "tobp", language$, tobp.Caption, tobp, default_font)
    RunPassFail.Caption = get_thing("selecttest", "runpassfail", language$, RunPassFail.Caption, RunPassFail, default_font)
    StopOnFail.Caption = get_thing("selecttest", "stoponfail", language$, StopOnFail.Caption, StopOnFail, default_font)
    Label1.Caption = get_thing("selecttest", "label1", language$, Label1.Caption, Label1, default_font)
    Label2.Caption = get_thing("selecttest", "label2", language$, Label2.Caption, Label2, default_font)
    set_fontstuff PListFilename, default_font
    PListCheck.Caption = get_thing("selecttest", "plistcheck", language$, PListCheck.Caption, PListCheck, default_font)
    singlePointCheck.Caption = get_thing("selecttest", "singlepoint", language$, singlePointCheck.Caption, singlePointCheck, default_font)
    set_fontname singlePointSetup, default_font
    singlePointSetup.Caption = gpps2("selecttest", "singlepointsetup", language$, singlePointSetup.Caption)
    
    ' Other text
    ts$(1) = gpps2("selecttest", "ts1", language$, "liq. vapor perm.")
    ts$(2) = gpps2("selecttest", "ts2", language$, "Unit")
    ts$(3) = gpps2("selecttest", "ts3", language$, "Illegal pass/fail values. Please reenter.")
    ts$(4) = gpps2("selecttest", "ts4", language$, "Choose pressure list")
    
End Sub

