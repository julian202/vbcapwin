VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form autoparm 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capillary Flow Autotesting Parameters"
   ClientHeight    =   9540
   ClientLeft      =   7605
   ClientTop       =   1110
   ClientWidth     =   8790
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
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9540
   ScaleWidth      =   8790
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Lohm Calibration"
      Height          =   9255
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Width           =   8535
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   22
         Left            =   3960
         TabIndex        =   75
         Top             =   6600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000080FF&
         Caption         =   "Enable Temperature Control"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   6360
         Width           =   2895
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   21
         Left            =   240
         TabIndex        =   73
         Top             =   6720
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   5550
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   10
         Left            =   3960
         TabIndex        =   28
         Top             =   1950
         Width           =   1095
      End
      Begin VB.CommandButton Parm_Help 
         Caption         =   "&Help"
         Height          =   495
         Left            =   6960
         TabIndex        =   0
         Top             =   870
         Width           =   1365
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   3390
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   11
         Left            =   3960
         TabIndex        =   30
         Top             =   2310
         Width           =   1095
      End
      Begin VB.CommandButton SaveDef 
         Caption         =   "Make &Default"
         Height          =   735
         Left            =   6960
         TabIndex        =   1
         Top             =   1470
         Width           =   1365
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   12
         Left            =   3960
         TabIndex        =   32
         Top             =   3030
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1590
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   13
         Left            =   3960
         TabIndex        =   34
         Top             =   3390
         Width           =   1095
      End
      Begin VB.CommandButton Loadnew 
         Caption         =   "&Load..."
         Height          =   615
         Left            =   6960
         TabIndex        =   2
         Top             =   2310
         Width           =   1365
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   2670
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   14
         Left            =   3960
         TabIndex        =   36
         Top             =   3750
         Width           =   1095
      End
      Begin VB.CommandButton Saveas 
         Caption         =   "Save &As..."
         Height          =   615
         Left            =   6960
         TabIndex        =   3
         Top             =   3030
         Width           =   1365
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   6
         Left            =   255
         TabIndex        =   8
         Top             =   5190
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   15
         Left            =   3960
         TabIndex        =   38
         Top             =   4710
         Width           =   1095
      End
      Begin VB.CommandButton printparms 
         Caption         =   "&Print"
         Height          =   615
         Left            =   6960
         TabIndex        =   5
         Top             =   4470
         Width           =   1365
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   870
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   16
         Left            =   3960
         TabIndex        =   40
         Top             =   5070
         Width           =   1095
      End
      Begin VB.CommandButton quitparm 
         Caption         =   "E&xit"
         Height          =   615
         Left            =   6960
         TabIndex        =   6
         Top             =   5190
         Width           =   1365
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   8
         Left            =   3960
         TabIndex        =   24
         Top             =   1230
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   52
         Top             =   4470
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   9
         Left            =   3960
         TabIndex        =   26
         Top             =   1605
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   18
         Left            =   240
         TabIndex        =   50
         Top             =   4110
         Width           =   1095
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   19
         Left            =   240
         TabIndex        =   18
         Top             =   2310
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   615
         Left            =   6960
         TabIndex        =   4
         Top             =   3750
         Width           =   1365
      End
      Begin VB.TextBox P_Value 
         Height          =   285
         Index           =   20
         Left            =   3960
         TabIndex        =   42
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdRSet 
         Caption         =   "Reset Defaults"
         Height          =   375
         Left            =   2880
         TabIndex        =   47
         Top             =   8640
         Width           =   2775
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   495
         Left            =   240
         TabIndex        =   44
         Top             =   7710
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   327682
         Max             =   100
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider Slider3 
         Height          =   495
         Left            =   5760
         TabIndex        =   45
         Top             =   7710
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   327682
         Max             =   100
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   495
         Left            =   3000
         TabIndex        =   46
         Top             =   7710
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   327682
         Max             =   100
         TickFrequency   =   10
      End
      Begin VB.Label wettingTimeLabel1 
         BackColor       =   &H000080FF&
         Caption         =   "seconds"
         Height          =   255
         Left            =   5160
         TabIndex        =   77
         Top             =   6600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label wettingTimeLabel 
         BackColor       =   &H000080FF&
         Caption         =   "Wetting Time:"
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
         Left            =   3960
         TabIndex        =   76
         Top             =   6240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "Celsius"
         Height          =   255
         Index           =   20
         Left            =   1440
         TabIndex        =   72
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H000080FF&
         Caption         =   "Temperature Control"
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
         Left            =   240
         TabIndex        =   71
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Label Parmsfilename 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "this filename"
         Height          =   600
         Left            =   285
         TabIndex        =   48
         Top             =   120
         Width           =   8010
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "pulsewidth (sec)"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   5550
         Width           =   1815
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "flowslew (cts)"
         Height          =   195
         Index           =   10
         Left            =   5160
         TabIndex        =   27
         Top             =   1950
         Width           =   1545
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "v2incr (cts)"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   11
         Top             =   3390
         Width           =   1815
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "eqiter (0.1 sec)"
         Height          =   195
         Index           =   11
         Left            =   5160
         TabIndex        =   29
         Top             =   2310
         Width           =   1545
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "F/PT (old bubltime)"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   15
         Top             =   1950
         Width           =   1815
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "aveiter (0.1 sec)"
         Height          =   195
         Index           =   12
         Left            =   5160
         TabIndex        =   31
         Top             =   3030
         Width           =   1500
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "maxpdif"
         Height          =   195
         Index           =   13
         Left            =   5160
         TabIndex        =   33
         Top             =   3390
         Width           =   1500
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "bublflow (cc/m)"
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   13
         Top             =   1590
         Width           =   1815
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "maxfdif (cc/m)"
         Height          =   195
         Index           =   14
         Left            =   5160
         TabIndex        =   35
         Top             =   3750
         Width           =   1500
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "zerotime (sec)"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   19
         Top             =   2670
         Width           =   1815
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "maxpres"
         Height          =   195
         Index           =   6
         Left            =   1440
         TabIndex        =   7
         Top             =   5190
         Width           =   1815
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "startp"
         Height          =   195
         Index           =   15
         Left            =   5160
         TabIndex        =   37
         Top             =   4710
         Width           =   1500
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "startf (cc/m)"
         Height          =   195
         Index           =   16
         Left            =   5160
         TabIndex        =   39
         Top             =   5070
         Width           =   1500
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "maxflow (cc/m)"
         Height          =   195
         Index           =   7
         Left            =   1440
         TabIndex        =   21
         Top             =   870
         Width           =   1815
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "pulse delay (sec)"
         Height          =   195
         Index           =   17
         Left            =   1440
         TabIndex        =   53
         Top             =   4470
         Width           =   1500
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "mineqtime (sec)"
         Height          =   195
         Index           =   8
         Left            =   5160
         TabIndex        =   23
         Top             =   1230
         Width           =   1575
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "preginc (cts)"
         Height          =   195
         Index           =   18
         Left            =   1440
         TabIndex        =   51
         Top             =   4110
         Width           =   1500
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "presslew (cts)"
         Height          =   195
         Index           =   9
         Left            =   5160
         TabIndex        =   25
         Top             =   1590
         Width           =   1575
      End
      Begin VB.Label MinMax 
         BackColor       =   &H80000005&
         Height          =   195
         Left            =   1560
         TabIndex        =   43
         Top             =   7110
         Width           =   5370
      End
      Begin VB.Label P_Value_Label 
         BackColor       =   &H000080FF&
         Caption         =   "minbppres"
         Height          =   195
         Index           =   19
         Left            =   1440
         TabIndex        =   17
         Top             =   2310
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Lohm Calibration"
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
         Left            =   240
         TabIndex        =   70
         Top             =   4830
         Width           =   3615
      End
      Begin VB.Label Stability1 
         BackColor       =   &H000080FF&
         Caption         =   "Stability Routine #1"
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
         Left            =   3960
         TabIndex        =   69
         Top             =   870
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Stability Routine #2"
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
         Left            =   3960
         TabIndex        =   68
         Top             =   2670
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         Caption         =   "Current Test Status Graph Scale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   67
         Top             =   4110
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         Caption         =   "Bubble Point/Integrity Test"
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
         Left            =   240
         TabIndex        =   66
         Top             =   1230
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackColor       =   &H000080FF&
         Caption         =   "Motorized Valve 2 Control"
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
         Left            =   240
         TabIndex        =   65
         Top             =   3030
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackColor       =   &H000080FF&
         Caption         =   "Regulator Control"
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
         Left            =   240
         TabIndex        =   64
         Top             =   3750
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackColor       =   &H000080FF&
         Caption         =   "Leak Test"
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
         Left            =   3960
         TabIndex        =   63
         Top             =   5400
         Width           =   2895
      End
      Begin VB.Label leaktest_wait 
         BackColor       =   &H000080FF&
         Caption         =   "read_delay (sec)"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   41
         Top             =   5790
         Width           =   1740
      End
      Begin VB.Label lblBubPt 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Bubble Point Test"
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   7470
         Width           =   2175
      End
      Begin VB.Label lblWetDryTest 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Wet/Dry Test"
         Height          =   255
         Left            =   3360
         TabIndex        =   61
         Top             =   7470
         Width           =   1935
      End
      Begin VB.Label lblStability 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Stability Test"
         Height          =   255
         Left            =   6120
         TabIndex        =   60
         Top             =   7470
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H000080FF&
         Caption         =   "Accurate"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   59
         Top             =   8190
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Fast"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   58
         Top             =   8190
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H000080FF&
         Caption         =   "Accurate"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   57
         Top             =   8190
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Fast"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   56
         Top             =   8190
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Accurate"
         Height          =   255
         Index           =   4
         Left            =   6840
         TabIndex        =   55
         Top             =   8190
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H000080FF&
         Caption         =   "Fast"
         Height          =   255
         Index           =   5
         Left            =   5760
         TabIndex        =   54
         Top             =   8190
         Width           =   975
      End
   End
End
Attribute VB_Name = "autoparm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim DryFile$
Dim WetFile$
Dim EditWhich$
Dim filename$
Dim value(MAXVALS%) As Single
Dim MIN(MAXVALS%) As Single
Dim max(MAXVALS%) As Single
Dim xform$(MAXVALS%)
Dim dirty As Integer
Dim ts$(29)                     ' Text strings for this form
Dim strO As String
Dim valueChanged As Boolean
'A.Bentley - change leak_test_delay range from 0 to 10 seconds to 0 to 300 seconds per Josh.


' E.Corvinelli May 2004
'This sub is to restore the P_Values to (4,3,2,18,11)if they are altered by the user with either
'the slider or manually inserting the value.
'P_Value_Change() gives the Textbox focus while P_Value_LostFocusmakes the program recognize it as
'the new value of that parameter.

Private Sub cmdRSet_Click()
Dim s As Single
P_Value_Change (4)
'   11-06-07    Changed the manuall inserting value
'P_Value(4).Text = Int(20)       'P_Value(4)buble flow, Slider 1 buble point test
P_Value(4).Text = Int(10)       'P_Value(4)buble flow, Slider 1 buble point test

P_Value_LostFocus (4)
P_Value_Change (3)
P_Value(3).Text = Int(50)       'P_Value(3)is F\PT, Slider 1 buble point test
P_Value_LostFocus (3)
P_Value_Change (2)
P_Value(2).Text = Int(10)       'P_Value(2) is v2incr, Slider2 Wet\Dry test
P_Value_LostFocus (2)
P_Value_Change (18)
P_Value(18).Text = Int(1)       'P_Value(18) is preginc, Slider2 Wet\Dry Test
P_Value_LostFocus (18)
P_Value_Change (11)
P_Value(11).Text = Int(20)      'P_Value(11) is equiter,Slider3 Stability Test
P_Value_LostFocus (11)
strO = P_Value(11).Text         'resets the slider position to corospond to the textbox value
s = val(strO)
If s < 0 Then s = 0
If s > 100 Then s = 100
Slider3.SelStart = (Int(s))
strO = P_Value(4).Text
s = val(strO) / 30 * 100
If s < 0 Then s = 0
If s > 100 Then s = 100
Slider1.SelStart = (Int(s))
strO = P_Value(2).Text
s = val(strO) / 5
If s < 0 Then s = 0
If s > 100 Then s = 100
Slider2.SelStart = (Int(s))
SaveDef.value = True            ' Writes to default file
'

End Sub

Private Sub cmdSave_Click()

    Dim filenum As Integer
    Dim i%
    
    fsel_return$ = LTrim$(RTrim$(Parmsfilename.Caption))
    file$ = fsel_return$
    filenum = FreeFile
    Open file$ For Output As #filenum
        For i% = 1 To MAXVALS%: Print #filenum, str$(value(i%)): Next i%
        Print #filenum, IIf(Check1.value, 1, 0)
    Close #filenum
    dirty = False
    Parmsfilename.Caption = UCase$(file$)
    MsgBox LCase$(file$) & " " + ts$(1)             ' "was saved."

End Sub

Private Sub Form_Load()
'The line below added by E.Cocvinelli 5\26\04 to display frmFastTest on
'the same screen @the same time as autoparm form.



Dim rep%, i%
Dim fn As Integer
Dim s As Single
valueChanged = False
LoadTextStrings

' Superseded by LoadTextStrings:
'Parmsfilename.FontName = system_font
'Parmsfilename.fontsize = font_size
'Parmsfilename.fontbold = font_bold
If version >= 7 Then
    P_Value_Label(2) = "v2incr (cts*3)"
    P_Value_Label(18) = "preginc (cts*50)"
    P_Value_Label(9) = "presslew (cts*3)"
    P_Value_Label(10) = "flowslew (cts*3)"
End If
    TitleScrn.modifymenu(0).Enabled = False
    If liqpermonly Then
        Caption = ts$(29) + " " + ts$(16)                    ' "Liquid Permeameter Autotesting Parameters"
        Label4.Visible = False
        Label5.Visible = False
        hide_unused_parameters (3)
        hide_unused_parameters (4)
        hide_unused_parameters (19)
        hide_unused_parameters (5)
        hide_unused_parameters (7)
        hide_unused_parameters (2)
        hide_unused_parameters (14) 'added edc 06-15-07
    ElseIf GasPerm Then
        Caption = ts$(2) + " " + ts$(16)                    ' "Gas Permeameter Autotesting Parameters"
        hide_unused_parameters (3)
        hide_unused_parameters (4)
        hide_unused_parameters (19)
        hide_unused_parameters (5)
    ElseIf CFAnal Then
        Caption = ts$(3) + " " + ts$(16)                    ' "CFA Autotesting Parameters"
    ElseIf BPTester Then
        ' 6.71.38x begin
        If hydrohead_exclusive Then
            Caption = ts$(26)                               ' "Hydrohead Tester Parameters"
        Else
            Caption = ts$(4) + " " + ts$(16)                    ' "Bubble Point Tester Parameters"
        End If
        ' 6.71.38x end
        hide_unused_parameters (1)
        hide_unused_parameters (2)
        hide_unused_parameters (8)
        hide_unused_parameters (9)
        hide_unused_parameters (10)
        hide_unused_parameters (11)
        hide_unused_parameters (7)
        hide_unused_parameters (12)
        hide_unused_parameters (13)
        hide_unused_parameters (14)
        Label5.Visible = False
        Stability1.Visible = False
        Label2.Visible = False
    ElseIf itester Then
        Caption = ts$(5) + " " + ts$(16)                ' "Integrity Tester Autotesting Parameters"
        hide_unused_parameters (1)
        hide_unused_parameters (2)
        hide_unused_parameters (6)
        hide_unused_parameters (7)
    Else
        Caption = ts$(6) + " " + ts$(16)                ' "Capillary Flow Autotesting Parameters"
    End If
   move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    'RemoveSysMenu Me
    Rem MIN(1) WAS 0.05
    
    MIN(1) = 0: max(1) = 1: xform$(1) = "##0.0000"
    ' changing min values
    MIN(2) = 1: max(2) = 500: xform$(2) = "#####0.#"
    'edc 05-22-07 F/PT min altered to 50
    'AJB 11-05-09 F/PT min altered to user defined per Dr. G
    MIN(3) = minFPT: max(3) = 999999: xform$(3) = "#####0.#"
    ' edc 05-22-07 bubble point minimum for lo pressure machine is 2 hi press min should be 10
    ' rvw 11-20-07 new capwin.ini parameter minbubflow with default set using similar logic
    'If dualregulator = True Then
    '    MIN(4) = 10#: max(4) = fy2(1, 0) * gasflowconversionfactor: xform$(4) = "##0.00"
    '    'edc 9-18-07 sets min to 20 if there is only one transducer and it is high pressure
    '    '11-6-07 Change the dualregulator from (PY2(0) >= 110) to (PY2(0) >= 160) for variable presure readings that might come up
    'ElseIf dualregulator = False And (PY2(0) >= 160) Then
    '    MIN(4) = 10#: max(4) = fy2(1, 0) * gasflowconversionfactor: xform$(4) = "##0.00"
    'Else
    '    MIN(4) = 5#: max(4) = fy2(1, 0) * gasflowconversionfactor: xform$(4) = "##0.00" ' min changed from 2 to 5 for .106
    'End If
    MIN(4) = minbubflow: max(4) = FY2(1, 0) * gasflowconversionfactor: xform$(4) = "##0.00"
    If max(4) < FY2(0, 0) * gasflowconversionfactor Then max(4) = FY2(0, 0) * gasflowconversionfactor
    MIN(5) = 0.5: max(5) = 30: xform$(5) = "###0.0"
    Rem min(6) was 1
    MIN(6) = 0: max(6) = PY2(0):  xform$(6) = "#####0.######"
    MIN(7) = 1: max(7) = FY2(1, 0) * gasflowconversionfactor: xform$(7) = "#####0"
    If FY2(1, 2) > FY2(1, 0) Then
        If hasMultipleMVs Then
            max(7) = FY2(2, 2) * gasflowconversionfactor
        Else
            max(7) = FY2(1, 2) * gasflowconversionfactor
        End If
    End If
    MIN(8) = 0: max(8) = 99999: xform$(8) = "#####0.#"
    MIN(9) = 1: max(9) = 100:  xform$(9) = "#####0.#"
    MIN(10) = 1: max(10) = 100:  xform$(10) = "#####0.#"
    Rem min(11) was 1
    MIN(11) = 0: max(11) = 100:  xform$(11) = "#####0.#"
    Rem MIN(12) WAS 2
    MIN(12) = 0: max(12) = 100:  xform$(12) = "#####0.#"
    MIN(13) = 0.01: max(13) = 10: xform$(13) = "##0.00"
    MIN(14) = 0.1: max(14) = FY2(1, 0) * gasflowconversionfactor: xform$(14) = "###0.0"
    MIN(15) = 0.01: max(15) = PY2(0): xform$(15) = "###0.0"
    'MIN(16) = 10: max(16) = FY2(1, 0) * gasflowconversionfactor: xform$(16) = "#####0"
    MIN(16) = 0.1: max(16) = FY2(1, 0) * gasflowconversionfactor: xform$(16) = "#####0"
    
    MIN(17) = 0: max(17) = 10: xform$(17) = "##0"
    MIN(18) = 1: max(18) = 30: xform$(18) = "#0.##"
    If version >= 7 Then MIN(18) = 1 / 50
    MIN(19) = 0: max(19) = PY2(0): xform$(19) = "###0.000"
    MIN(20) = 0: max(20) = 300: xform(20) = "##0.00" 'AJB
    MIN(21) = 0: max(21) = 200: xform(21) = "##0.0" 'AJB 10-20-09
    MIN(22) = 0: max(22) = 999: xform(22) = "0.0" 'ajb 12-24-09
    
    P_Value_Label(6).Caption = "maxpres (" + PU$ + ")"
    P_Value_Label(13).Caption = "maxpdif (" + PU$ + ")"
    P_Value_Label(15).Caption = "startp (" + PU$ + ")"
    P_Value_Label(19).Caption = "minbppres (" + PU$ + ")"
                  
    ' set default file name for editing
    filename$ = EXE_Path$ + "users\" + Curr_U$ + "\DEFAULT.TPF"
    ' this may be modified by this section:
    If TestScreenVisible Then
        'are wet and dry files both being shown
        If Testscrn.testval(8).Visible = True And Testscrn.testval(9).Visible = True Then
            'if both data files are the same, choose one to open
            DryFile$ = TPFDRY$(current_unit%)
            WetFile$ = TPFWET$(current_unit%)
            If DryFile$ = WetFile$ Then
                filename$ = DryFile$
            Else
                ' don't ask if they are already on one of the lines
                If val(Testscrn.Tag) = 8 Then
                    filename$ = DryFile$
                    EditWhich$ = "DRY"
                ElseIf val(Testscrn.Tag) = 9 Then
                    filename$ = WetFile$
                    EditWhich$ = "WET"
                Else
                    'if the wet and dry curves are different files, then ask the user
                    'which one to edit
                    rep% = MsgBox(ts$(7), 4, ts$(8))     ' "Would you like to edit the Dry Curve? (Select no to edit             Wet Curve)"/"Choose File to Edit"
                    'yes
                    If rep% = 6 Then
                        filename$ = DryFile$
                        EditWhich$ = "DRY"
                    'no
                    Else
                        filename$ = WetFile$
                        EditWhich$ = "WET"
                    End If
                End If
            End If
        'if 9 is shown, it is guaranteed to be a file name
        ElseIf Testscrn.testval(9).Visible = True Then
            WetFile$ = TPFWET$(current_unit%)
            filename$ = WetFile$
        'if 8 is showing, check to make sure that it has not been changed to handle
        'the diff perm test.  If it doesn't say "Pressure...blah blah blah", then
        'it should be a filename
        ElseIf LCase$(Left$(Testscrn.testval(8).Caption, 8)) <> "pressure" Then
            DryFile$ = TPFDRY$(current_unit%)
            filename$ = DryFile$
        End If
    End If

    Parmsfilename.Caption = UCase$(filename$)
    fn = FreeFile
    Open filename$ For Input As #fn
        For i% = 1 To MAXVALS%
            On Error Resume Next
            value(i%) = 0
            Input #fn, value(i%)
            On Error GoTo 0
            If liqpermonly = True And i% = 6 Then
            value(i%) = max(i%)
            valueChanged = True
            End If
            If i% = 6 Or i% = 13 Or i% = 15 Or i% = 19 Then
                P_Value(i%).Text = Format$(value(i%) * PCNV, xform$(i%))
            Else
                P_Value(i%).Text = Format$(value(i%), xform$(i%))
            End If
            If (i% = 3 Or i% = 4 Or i% = 13 Or i% = 14 Or i% = 15 Or i% = 16 Or i% = 21) And P_Value(i%).Text < MIN(i%) Then
                    P_Value(i%).Text = MIN(i%)
                    value(i%) = myVal(P_Value(i%))
                    valueChanged = True
            End If
        Next i%
        On Error Resume Next
        Dim tempEnable As Integer
        Input #fn, tempEnable
        Check1.value = IIf(tempEnable, vbChecked, vbUnchecked)
    Close #fn
    If FY2(1, 2) > FY2(1, 0) Then max(16) = FY2(1, 2) * gasflowconversionfactor
    ' edc 06-04-07 code altered to allow the minimum and the parrameter to be displayed
    'and the default file or parameter file is altered in the event of this oddity this
    'will prevent future errors
    If max(16) < MIN(16) Then
    max(16) = MIN(16) + 1
    P_Value(16).Text = MIN(16)
         Open filename$ For Output As #fn
        For i% = 1 To MAXVALS%: Print #fn, str$(value(i%)): Next i%
    End If
    Close #fn
    '06-04-07 end of alteration
    If valueChanged = True Then
        Open filename$ For Output As #fn
        For i% = 1 To MAXVALS%: Print #fn, str$(value(i%)): Next i%
    Close #fn
    valueChanged = False
    End If
    path(1) = filename$
    Do Until Right$(path(1), 1) = "\"
        path(1) = Left(path(1), Len(path(1)) - 1)
    Loop
    dirty = False
    path(1) = path(1) & "*.tpf"
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(9) + str$(unitnumber)         ' "Unit"
    End If
    'EDC 12-08-04
    'next two lines added to start Slider3 at the value that is specified by the
    'by the text box P_Value(11)
    strO = P_Value(11).Text
    s = val(strO)
    If s < 0 Then s = 0
    If s > 100 Then s = 100
    Slider3.SelStart = (Int(s))
    strO = P_Value(4).Text
    s = val(strO) / 30 * 100
    If s < 0 Then s = 0
    If s > 100 Then s = 100
    Slider1.SelStart = (Int(s))
    strO = P_Value(2).Text
    s = val(strO) / 5
    If s < 0 Then s = 0
    If s > 100 Then s = 100
    Slider2.SelStart = (Int(s))
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
    'ajb 12-24-09
    If number_of_wetting_valves > 0 Then
        wettingTimeLabel.Visible = True
        wettingTimeLabel1.Visible = True
        P_Value(22).Visible = True
    End If
End Sub


Private Sub Form_Unload(cancel As Integer)

    Dim i%, fn As Integer
    
    
      reply% = 999
    If dirty Then reply% = MsgBox(ts$(10) + " " + filename$ + " " + ts$(11) + vbCrLf + ts$(12), 35, ts$(13)) '"Changes to"/"have not been saved."/"Would you like to save before closing?"/"File Not Saved"
    If reply% = 2 Then cancel = -1
    If reply% = 6 Then
        fsel_name$ = filename$
        fsel_title$ = ts$(6) + " " + ts$(14)                             ' "Capillary Flow Parameter File"
        fsel_path$ = path(1)
        fsel_io = False
        fsel Me.hwnd
        If fsel_return$ = "" Then
            Exit Sub
        End If
        path(1) = fsel_path$
        file$ = fsel_return$
        fn = FreeFile
        Open file$ For Output As #fn
            For i% = 1 To MAXVALS%
                Print #fn, str$(value(i%))
            Next i%
        Close #fn
    End If
    If TestScreenVisible Then
        'are wet and dry files both being shown
        If Testscrn.testval(8).Visible = True And Testscrn.testval(9).Visible = True Then
            If DryFile$ = WetFile$ Then
                'if both data files are the same, change label of both
                TPFDRY$(current_unit%) = UCase$(filename$)
                TPFWET$(current_unit%) = TPFDRY$(current_unit%)
                Testscrn.testval(8).Caption = Testscrn.make_sure(shorten(TPFDRY$(current_unit%)))
                Testscrn.testval(9).Caption = Testscrn.make_sure(shorten(TPFWET$(current_unit%)))
            Else
                If EditWhich$ = "DRY" Then
                    TPFDRY$(current_unit%) = UCase$(filename$)
                    Testscrn.testval(8).Caption = Testscrn.make_sure(shorten(TPFDRY$(current_unit%)))
                Else
                    TPFWET$(current_unit%) = UCase$(filename$)
                    Testscrn.testval(9).Caption = Testscrn.make_sure(shorten(TPFWET$(current_unit%)))
                End If
            End If
        'if 9 is shown, it is guaranteed to be a file name
        ElseIf Testscrn.testval(9).Visible = True Then
            TPFWET$(current_unit%) = UCase$(filename$)
            Testscrn.testval(9).Caption = Testscrn.make_sure(shorten(TPFWET$(current_unit%)))
        'if 8 is showing, check to make sure that it has not been changed to handle
        'the diff perm test.  If it doesn't say "Pressure...blah blah blah", then
        'it should be a filename
        ElseIf LCase$(Left$(Testscrn.testval(8).Caption, 8)) <> "pressure" Then
            TPFDRY$(current_unit%) = UCase$(filename$)
            Testscrn.testval(8).Caption = Testscrn.make_sure(shorten(TPFDRY$(current_unit%)))
        End If
    End If
    If (cancel <> -1) Then TitleScrn.modifymenu(0).Enabled = True

End Sub

Private Sub hide_unused_parameters(i%)
   P_Value(i%).Visible = False
   P_Value_Label(i%).Visible = False
End Sub

Private Sub Loadnew_Click()

    Dim i%, fn As Integer
    
    If dirty Then
        reply% = MsgBox(ts$(10) + " " + filename$ + vbCrLf + " " + ts$(11) + vbCrLf + ts$(15), 1, ts$(13)) ' "Changes to"/"have not been saved."/"Select OK to Load New Parameters."/"File not saved"
        If reply% = 1 Then dirty = False Else Exit Sub
    End If
    
    fsel_name$ = ""
    If GasPerm Then
        fsel_title$ = ts$(2) + " " + ts$(14)        ' "Gas Permeameter Parameter File"
    ElseIf CFAnal Then
        fsel_title$ = ts$(3) + " " + ts$(14)        ' "CFA Parameter File"
    ElseIf itester Then
        fsel_title$ = ts$(5) + " " + ts$(14)        ' "Integrity Tester Parameter File"
    ElseIf BPTester Then
        ' 6.71.38x begin
        If hydrohead_exclusive Then
            fsel_title$ = ts$(27)                   ' "Hydrohead Tester Parameter File"
        Else
            fsel_title$ = ts$(4) + " " + ts$(14)        ' "Bubble Point Tester Parameter File"
        End If
        ' 6.71.38x end
    Else
        fsel_title$ = ts$(6) + " " + ts$(14)        ' "Capillary Flow Parameter File"
    End If
    fsel_path$ = path(1)
    fsel_io = True
    fsel Me.hwnd
    If fsel_return$ = "" Then
        Exit Sub
    End If
    path(1) = fsel_path$
    file$ = fsel_return$
    
    fn = FreeFile
    Open file$ For Input As #fn
        For i% = 1 To MAXVALS%
            On Error Resume Next
            value(i%) = 0
            Input #fn, value(i%)
            On Error GoTo 0
            If i% = 6 Or i% = 13 Or i% = 15 Or i% = 19 Then
                P_Value(i%).Text = Format$(value(i%) * PCNV, xform$(i%))
            Else
                P_Value(i%).Text = Format$(value(i%), xform$(i%))
            End If
            'edc 05-22-07 check to see if F/PT or bubble point are within minimuns
            If (i% = 3 Or i% = 4 Or i% = 13 Or i% = 14 Or i% = 15 Or i% = 16) And P_Value(i%).Text < MIN(i%) Then
                    P_Value(i%).Text = MIN(i%)
                    value(i%) = myVal(P_Value(i%))
                    valueChanged = True
            End If
        Next i%
    Close #fn
    If valueChanged = True Then
        Open file$ For Output As #fn
        For i% = 1 To MAXVALS%: Print #fn, str$(value(i%)): Next i%
    Close #fn
    End If
    valueChanged = False
    filename$ = file$
    Parmsfilename.Caption = UCase$(filename$)
    dirty = False
End Sub

Private Sub Loadnew_GotFocus()
    MinMax.Caption = ""
End Sub

Private Sub P_Value_Change(Index As Integer)
    dirty = True
End Sub
'altered the formatting of the minimum string so the max was viewable in all cases
Private Sub P_Value_GotFocus(Index As Integer)
    Dim temp As Single
    P_Value(Index).SelStart = 0
    P_Value(Index).SelLength = 10
    
    If Index = 6 Or Index = 13 Or Index = 15 Or Index = 19 Then
        temp = Format((MIN(Index) * PCNV), "##0.00000")
        MinMax.Caption = ts$(17) + " = " + str$(temp) + " " + ts$(18) + " = " + str$(max(Index) * PCNV) '"Current Parameter Min"/"and Max"
    Else
        temp = Format((MIN(Index) * PCNV), "##0.00000")
       MinMax.Caption = ts$(17) + " = " + str$(temp) + " " + ts$(18) + " = " + str$(max(Index))              '  "Current Parameter Min"/"and Max"
    End If
    
End Sub

Private Sub P_Value_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case KeyAscii
        ' allow comma (44) as well as decimal point (46)
        Case 8, Asc(Decimal_Point$), 48 To 57
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub P_Value_LostFocus(Index As Integer)

    Dim ConversionX As Single
    
    If Index = 6 Or Index = 13 Or Index = 15 Or Index = 19 Then
        ConversionX = PCNV
    Else
        ConversionX = 1
    End If
    If (myVal(P_Value(Index).Text) < MIN(Index) * ConversionX - 0.001) Or (myVal(P_Value(Index).Text) > max(Index) * ConversionX + 0.001) Then
        Beep
        MsgBox ts$(19)              ' ("Entered Parameter Out of Allowable Range - See Message Box Below For Legal Entries")
        If Index = 3 Then P_Value(Index).Text = MIN(Index)
        If Index = 4 Then P_Value(Index).Text = MIN(Index)
        If TypeOf Screen.ActiveControl Is textbox Then P_Value(Index).SetFocus
        Exit Sub
    Else
        value(Index) = myVal(P_Value(Index).Text) / ConversionX
    End If

End Sub

Private Sub Parm_Help_Click()
    
    Dim r As Long
    
    r = WinHelp(hwnd, HelpFile$, Help_Context, ByVal 20&)

End Sub

Private Sub Parm_Help_GotFocus()
    MinMax.Caption = ""
End Sub

Private Sub printparms_Click()
' Print the current parameter values

    Dim temp$, i%, j%, col1(10) As Integer, col2(10) As Integer
    Dim padding$

    Rem don't you miss the old DATA statement at these times?
    col1(1) = 7
    col1(2) = 4
    col1(3) = 3
    col1(4) = 19
    col1(5) = 5
    col1(6) = 2
    col1(7) = 18
    col1(8) = 17
    col1(9) = 6
    col1(10) = 1
    col2(1) = 8
    col2(2) = 9
    col2(3) = 10
    col2(4) = 11
    col2(5) = 12
    col2(6) = 13
    col2(7) = 14
    col2(8) = 15
    col2(9) = 16
    col2(10) = -1

    On Error GoTo No_Printer
    Printer.fontsize = 12

    ' Calculate padding for centering of filename
    If Len(temp$) > 80 Then
        padding$ = ""
    Else
        padding$ = String$(40 - Len(temp$) / 2, " ")
    End If
    Printer.Print vbCrLf + vbCrLf
    temp$ = ts$(6) + " " + ts$(16)  ' "Capillary Flow AutoTesting Parameters"
    Printer.Print padding$ + temp$ + vbCrLf + vbCrLf
    temp$ = UCase$(filename$)
    Printer.Print padding$ + temp$ + vbCrLf + vbCrLf
    For i% = 1 To 10
        j% = col1(i%)
        Printer.Print Tab(3);
        Printer.Print P_Value(j%).Text;
        Printer.Print Tab(13);
        Printer.Print P_Value_Label(j%).Caption;
        If i% < 10 Then
            Printer.Print Tab(40);
            Printer.Print P_Value(col2(i%)).Text;
            Printer.Print Tab(50);
            Printer.Print P_Value_Label(col2(i%)).Caption;
        End If
        Printer.Print
    Next i%
    Printer.EndDoc
    On Error GoTo 0
    Exit Sub
No_Printer:
    MsgBox ts$(20), 16, ts$(21)         ' "There is a problem with the printer."/"Printer Error"
    Resume P_Finis
P_Finis:
    On Error GoTo 0
    
End Sub

Private Sub printparms_GotFocus()
    MinMax.Caption = ""
End Sub

Private Sub quitparm_Click()
    Unload Me
End Sub

Private Sub quitparm_GotFocus()
    MinMax.Caption = ""
End Sub

Private Sub Saveas_Click()

    Dim i%, fn As Integer
    
    fsel_name$ = LTrim$(RTrim$(Parmsfilename.Caption))
    If GasPerm Then
        fsel_title$ = ts$(2) + " " + ts$(22)        ' "Gas Permeameter Parameter File Save AS."
    ElseIf CFAnal Then
        fsel_title$ = ts$(3) + " " + ts$(22)        ' "CFA Parameter File Save AS."
    ElseIf itester Then
        fsel_title$ = ts$(5) + " " + ts$(22)        ' "Integrity Tester Parameter File Save AS."
    ElseIf BPTester Then
        ' 6.71.38x begin
        If hydrohead_exclusive Then
            fsel_title$ = ts$(28)                   ' "Hydrohead Tester Parameter File Save AS"
        Else
            fsel_title$ = ts$(4) + " " + ts$(22)        ' "Bubble Point Tester Parameter File Save AS."
        End If
        ' 6.71.38x end
    Else
        fsel_title$ = ts$(6) + " " + ts$(22)        ' "Capillary Flow Parameter File Save AS."
    End If
    fsel_path$ = path(1)
    fsel_io = False
    fsel Me.hwnd
    If fsel_return$ = "" Then
        Exit Sub
    End If
    path(1) = fsel_path$
    file$ = fsel_return$
    filename$ = file$

    fn = FreeFile
    Open file$ For Output As #fn
        For i% = 1 To MAXVALS%
            If i = 3 And value(i) < minFPT Then
                value(i) = minFPT
                MsgBox ("F/PT is below " + str$(minFPT) + " this is an unacceptable value. F/PT has been saved as 0.")
                P_Value(i).Text = minFPT
            End If
            Print #fn, str$(value(i%))
        Next i%
    Close #fn
    dirty = False
    Parmsfilename.Caption = UCase$(file$)
    
End Sub

Private Sub Saveas_GotFocus()
    MinMax.Caption = ""
End Sub

Private Sub SaveDef_Click()

    Dim i%, fn As Integer
    
    reply% = MsgBox(ts$(23) + ".tpf" + " " + ts$(24), 4, ts$(25)) ' "Overwite the existing default"/"with the current values?"/"Confirm"
    If reply% = 7 Then Exit Sub
    fn = FreeFile
    Open EXE_Path$ + "users\" + Curr_U$ + "\default.tpf" For Output As #fn
        For i% = 1 To MAXVALS%
            Print #fn, str$(value(i%))
        Next i%
    Close #fn
    If dirty = True And UCase$(filename$) = UCase$(EXE_Path$ + "users\" + Curr_U$ + "\DEFAULT.TPF") Then
        dirty = False
    End If
    
End Sub

Private Sub SaveDef_GotFocus()
    MinMax.Caption = ""
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    autoparm.Caption = gpps2("autoparm", "window title", language$, autoparm.Caption)
    set_fontstuff Parmsfilename, default_font
    set_fontstuff MinMax, default_font
    Label1.Caption = get_thing("autoparm", "label1", language$, Label1.Caption, Label1, default_font)
    Label1.fontsize = Label1.fontsize + 2
    Label2.Caption = get_thing("autoparm", "label2", language$, Label2.Caption, Label2, default_font)
    Label2.fontsize = Label2.fontsize + 2
    Label3.Caption = get_thing("autoparm", "label3", language$, Label3.Caption, Label3, default_font)
    Label3.fontsize = Label3.fontsize + 2
    Label4.Caption = get_thing("autoparm", "label4", language$, Label4.Caption, Label4, default_font)
    Label4.fontsize = Label4.fontsize + 2
    Label5.Caption = get_thing("autoparm", "label5", language$, Label5.Caption, Label5, default_font)
    Label5.fontsize = Label5.fontsize + 2
    Label6.Caption = get_thing("autoparm", "label6", language$, Label6.Caption, Label6, default_font)
    Label6.fontsize = Label6.fontsize + 2
    Loadnew.Caption = gpps2("autoparm", "loadnew", language$, Loadnew.Caption)
    set_fontname Loadnew, default_font
    Parm_Help.Caption = gpps2("autoparm", "help", language$, Parm_Help.Caption)
    set_fontname Parm_Help, default_font
    printparms.Caption = gpps2("autoparm", "print", language$, printparms.Caption)
    set_fontname printparms, default_font
    quitparm.Caption = gpps2("autoparm", "quit", language$, quitparm.Caption)
    set_fontname quitparm, default_font
    Saveas.Caption = gpps2("autoparm", "saveas", language$, Saveas.Caption)
    set_fontname Saveas, default_font
    SaveDef.Caption = gpps2("autoparm", "makedef", language$, SaveDef.Caption)
    set_fontname SaveDef, default_font
    Stability1.Caption = get_thing("autoparm", "stability1", language$, Stability1.Caption, Stability1, default_font)
    Stability1.fontsize = Stability1.fontsize + 2
    cmdSave.Caption = gpps2("autoparm", "save", language$, cmdSave.Caption)
    set_fontname cmdSave, default_font
    For i = 1 To 19
        set_fontstuff P_Value_Label(i), default_font
    Next i
    Label7.Caption = get_thing("autoparm", "label7", language$, Label7.Caption, Label7, default_font)
    leaktest_wait(0).Caption = get_thing("autoparm", "leaktestwait", language$, leaktest_wait(0).Caption, leaktest_wait(0), default_font)
    
    lblBubPt.Caption = get_thing("autoparm", "lblbubpt", language$, lblBubPt.Caption, lblBubPt, default_font)
    lblWetDryTest.Caption = get_thing("autoparm", "lblwetdry", language$, lblWetDryTest.Caption, lblWetDryTest, default_font)
    lblStability.Caption = get_thing("autoparm", "lblstability", language$, lblStability.Caption, lblStability, default_font)
    For i = 0 To 5
        Label8(i).Caption = get_thing("autoparm", "label8" + Format$(i), language$, Label8(i).Caption, Label8(i), default_font)
    Next i
    set_fontname cmdRSet, default_font
    cmdRSet.Caption = gpps2("autoparm", "cmdrset", language$, cmdRSet.Caption)

    ' Other text
    ts$(1) = gpps2("autoparm", "ts1", language$, "was saved.")
    ts$(2) = gpps2("autoparm", "ts2", language$, "Gas Permeameter")
    ts$(3) = gpps2("autoparm", "ts3", language$, "CFA")
    ts$(4) = gpps2("autoparm", "ts4", language$, "Bubble Point Tester")
    ts$(5) = gpps2("autoparm", "ts5", language$, "Integrity Tester")
    ts$(6) = gpps2("autoparm", "ts6", language$, "Capillary Flow")
    ts$(7) = gpps2("autoparm", "ts7", language$, "Would you like to edit the Dry Curve? (Select no to edit Wet Curve)")
    ts$(8) = gpps2("autoparm", "ts8", language$, "Choose File to Edit")
    ts$(9) = gpps2("autoparm", "ts9", language$, "Unit")
    ts$(10) = gpps2("autoparm", "ts10", language$, "Changes to")
    ts$(11) = gpps2("autoparm", "ts11", language$, "have not been saved.")
    ts$(12) = gpps2("autoparm", "ts12", language$, "Would you like to save before closing?")
    ts$(13) = gpps2("autoparm", "ts13", language$, "File Not Saved")
    ts$(14) = gpps2("autoparm", "ts14", language$, "Parameter File")
    ts$(15) = gpps2("autoparm", "ts15", language$, "Select OK to load new parameters.")
    ts$(16) = gpps2("autoparm", "ts16", language$, "Autotesting Parameters")
    ts$(17) = gpps2("autoparm", "ts17", language$, "Current Parameter Min")
    ts$(18) = gpps2("autoparm", "ts18", language$, "and Max")
    ts$(19) = gpps2("autoparm", "ts19", language$, "Entered parameter out of allowable range - see message box below for legal entries")
    ts$(20) = gpps2("autoparm", "ts20", language$, "There is a problem with the printer.")
    ts$(21) = gpps2("autoparm", "ts21", language$, "Printer Error")
    ts$(22) = gpps2("autoparm", "ts22", language$, "File save as")
    ts$(23) = gpps2("autoparm", "ts23", language$, "Overwite the existing default")
    ts$(24) = gpps2("autoparm", "ts24", language$, "with the current values?")
    ts$(25) = gpps2("autoparm", "ts25", language$, "Confirm")
    If hhRunAsBurst(current_unit%) Then
        ts$(26) = "Burst Tester Data File"
        ts$(27) = "Burst Tester Parameter File"
        ts$(28) = "Burst Tester Parameter File Save AS"
    Else
        ts$(26) = gpps2("autoparm", "ts26", language$, "Hydrohead Tester Data File")
        ts$(27) = gpps2("autoparm", "ts27", language$, "Hydrohead Tester Parameter File")
        ts$(28) = gpps2("autoparm", "ts28", language$, "Hydrohead Tester Parameter File Save AS")
    End If
    ts$(29) = gpps2("autoparm", "ts29", language$, "Liquid Permeameter")
    
    
End Sub


'E.Corvinelli May 2004 This Subprogram invokes the P_Value_Change() subprogram to give the textboxes
'the vlaue of "dirty = True"
'EdC 7-22-04 alteration to remove reference to F/PT P_Value(3)

Private Sub Slider1_Change()
P_Value_Change (4)
    'P_Value_Change (3)

End Sub

'E.Corvinelli May 2004. This subprogram invokes P_Value_LostFocus() for index 3and 4;this sets the new
'value of the textbox to the variable Value(index), this allows the the program to recognize the
'value in the textbox and write it to a given file before test execution.
'EdC 7-22-04 alteration to remove reference to F/PT P_Value(3)
Private Sub Slider1_LostFocus()
P_Value_LostFocus (4)
'P_Value_LostFocus (3)

End Sub


'E.Corvinelli May 2004. This subprogram is designed  to allow the user to alter the values of
'P_Value(4 and 3),bubleFlow and F\PT, with the use of a slider rather than entering the values
'manually. This is to speed up the elapsed time it takes to perfom a Buble Piont Test.
'EdC 7-22-04 alteration to remove reference to F/PT P_Value(3)
Private Sub Slider1_Scroll()
Dim b As Single
Slider1.MIN = 0              'sets MIN,MAX,SmallChange,LargeChange
    Slider1.max = 100
    Slider1.SmallChange = 1
    Slider1.LargeChange = 25
    
    
    If (Slider1.value = 0) Then          'values if slider is 0
        P_Value(4).Text = (0.05)
        'P_Value(3).Text = (5)
    ElseIf (Slider1.value <= 99) Then            'values if slider is 1-99
    b = Slider1.value * 0.3
            P_Value(4).Text = (b)
        'P_Value(3).Text = (Int(Slider1.Value * 99.9995))
    Else                                        'Values if slider nis at 100
        P_Value(4).Text = (Int(30))
        'P_Value(3).Text = (Int(9999))
    End If
    
End Sub

'E.Corvinelli May 2004 This Subprogram invokes the P_Value_Change() subprogram to give the textboxes
'the vlaue of "dirty = True"
Private Sub Slider2_Change()
P_Value_Change (2)
    P_Value_Change (18)

End Sub

'E.Corvinelli May 2004. This subprogram invokes P_Value_LostFocus() for index 3and 4;this sets the new
'value of the textbox to the variable Value(index), this allows the the program to recognize the
'value in the textbox and write it to a given file before test execution.
Private Sub Slider2_LostFocus()
P_Value_LostFocus (2)
P_Value_LostFocus (18)

End Sub

'E.Corvinelli May 2004. This subprogram is designed  to allow the user to alter the values of
'P_Value(2 and 18),v2incr and preginc, with the use of a slider rather than entering the values
'manually. This is to speed up the elapsed time it takes to perfom a Wet\Dry Test.

Private Sub Slider2_Scroll()
Slider2.MIN = 0              'sets MIN,MAX,SmallChange,LargeChange
    Slider2.max = 100
    Slider2.SmallChange = 1
    Slider2.LargeChange = 25
    If (Slider2.value = 0) Then          'values if slider is 0
        P_Value(2).Text = (Int(1))
        P_Value(18).Text = ".02"
    ElseIf (Slider2.value <= 99) Then            'values if slider is 1-99
        P_Value(2).Text = (Int(Slider2.value * 5))
        P_Value(18).Text = Slider2.value * 0.3
    Else                                        'Values if slider nis at 100
        P_Value(2).Text = (Int(500))
        P_Value(18).Text = (Int(30))
    End If
End Sub

'E.Corvinelli May 2004 This Subprogram invokes the P_Value_Change() subprogram to give the textboxes
'the vlaue of "dirty = True"
Private Sub Slider3_Change()
P_Value_Change (11)
    
End Sub

'E.Corvinelli May 2004. This subprogram invokes P_Value_LostFocus() for index 3and 4;this sets the new
'value of the textbox to the variable Value(index), this allows the the program to recognize the
'value in the textbox and write it to a given file before test execution.
Private Sub Slider3_LostFocus()
P_Value_LostFocus (11)

End Sub

'E.Corvinelli May 2004. This subprogram is designed  to allow the user to alter the values of
'P_Value(11),eqiter with the use of a slider rather than entering the values
'manually. This is to speed up the elapsed time it takes to perfom a Stability Test.

Private Sub Slider3_Scroll()
 Slider3.MIN = 0              'sets MIN,MAX,SmallChange,LargeChange
    Slider3.max = 100
    Slider3.SmallChange = 1
    Slider3.LargeChange = 25
    
    If (Slider3.value = 0) Then          'values if slider is 0
        P_Value(11).Text = (Int(0))
    ElseIf (Slider3.value <= 99) Then            'values if slider is 1-99
        P_Value(11).Text = (Int(Slider3.value * 1))
    Else                                        'Values if slider nis at 100
        P_Value(11).Text = (Int(100))
    End If
End Sub

