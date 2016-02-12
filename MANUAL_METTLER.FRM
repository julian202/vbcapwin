VERSION 5.00
Begin VB.Form ManualControl_Mettler 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Control"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8580
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   8580
   Begin VB.CommandButton Command5 
      Caption         =   "Open Door"
      Height          =   255
      Left            =   6000
      TabIndex        =   127
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Autocal"
      Height          =   255
      Left            =   2160
      TabIndex        =   122
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Bottom Level"
      Height          =   495
      Left            =   7440
      TabIndex        =   121
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Zero Level"
      Height          =   375
      Left            =   7440
      TabIndex        =   120
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   255
      Left            =   7560
      TabIndex        =   119
      Top             =   1080
      Width           =   855
   End
   Begin VB.HScrollBar cregscroll 
      Height          =   255
      LargeChange     =   10
      Left            =   6960
      Max             =   255
      Min             =   1
      TabIndex        =   73
      Top             =   3840
      Value           =   20
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton lfctrl 
      Caption         =   "LFLOW ZERO"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   71
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton lfctrl 
      Caption         =   "LFLOW DOWN"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   70
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton lfctrl 
      Caption         =   "LFLOW UP"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   69
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.HScrollBar lfcscroll 
      Height          =   255
      LargeChange     =   10
      Left            =   2280
      Max             =   255
      Min             =   1
      TabIndex        =   67
      Top             =   6120
      Value           =   20
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.HScrollBar RegScroll 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      Min             =   1
      TabIndex        =   65
      Top             =   5760
      Value           =   20
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton new_debug 
      Caption         =   "debug"
      Height          =   975
      Left            =   8280
      TabIndex        =   64
      Top             =   0
      Width           =   255
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "10"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   6600
      TabIndex        =   63
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "9"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   6120
      TabIndex        =   62
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "8"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   61
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "7"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   6120
      TabIndex        =   60
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "6"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   59
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "5"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   58
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "4"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   57
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame pen2frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   495
      Left            =   2520
      TabIndex        =   54
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
      Begin VB.OptionButton penselect 
         Caption         =   "Penetrometer 2"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   56
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton penselect 
         Caption         =   "Penetrometer 1"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CheckBox xhhvalve 
      Caption         =   "Vent"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   53
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox xhhvalve 
      Caption         =   "Fill"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   52
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox xhhvalve 
      Caption         =   "Press."
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   51
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame regselframe 
      BorderStyle     =   0  'None
      Caption         =   "regselframe"
      Height          =   255
      Left            =   240
      TabIndex        =   47
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton regsel 
         Caption         =   "Reg2"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   49
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton regsel 
         Caption         =   "Reg1"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton cregff 
      Height          =   375
      Left            =   8040
      Picture         =   "MANUAL_METTLER.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cregstop 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      Picture         =   "MANUAL_METTLER.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cregrew 
      Height          =   375
      Left            =   7320
      Picture         =   "MANUAL_METTLER.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton regff 
      Height          =   375
      Left            =   1200
      Picture         =   "MANUAL_METTLER.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton regstop 
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      Picture         =   "MANUAL_METTLER.frx":0408
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton regrew 
      Height          =   375
      Left            =   240
      Picture         =   "MANUAL_METTLER.frx":050A
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox readatcheck 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3600
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "3"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   32
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "2"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   31
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chamberiso 
      BackColor       =   &H000080FF&
      Caption         =   "1"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   30
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton tempprobe2 
      Caption         =   "Bottom Probe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   2400
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton tempprobe1 
      Caption         =   "Top Probe"
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
      TabIndex        =   28
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton tempset 
      Caption         =   "Set"
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
      Left            =   240
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox tempval 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   25
      Text            =   "0"
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   3720
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Liq. Drain"
      Height          =   255
      Index           =   17
      Left            =   6000
      TabIndex        =   128
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line vcloseline 
      Index           =   16
      Visible         =   0   'False
      X1              =   5640
      X2              =   5880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line vopenline 
      Index           =   16
      Visible         =   0   'False
      X1              =   5760
      X2              =   5760
      Y1              =   1680
      Y2              =   1920
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V22"
      Height          =   255
      Index           =   16
      Left            =   5880
      TabIndex        =   125
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vcloseline 
      Index           =   15
      Visible         =   0   'False
      X1              =   2520
      X2              =   2760
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line vopenline 
      Index           =   15
      Visible         =   0   'False
      X1              =   2640
      X2              =   2640
      Y1              =   2400
      Y2              =   2640
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V21"
      Height          =   255
      Index           =   15
      Left            =   2760
      TabIndex        =   123
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label pdebug 
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   118
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label pdebug 
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   117
      Top             =   1080
      Width           =   855
   End
   Begin VB.Line auxbottubing 
      Index           =   2
      Visible         =   0   'False
      X1              =   5640
      X2              =   4680
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label v2label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "V2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   116
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label pistonlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Piston"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   115
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape helpshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   360
      Width           =   255
   End
   Begin VB.Line auxtoptubing 
      Visible         =   0   'False
      X1              =   5880
      X2              =   5880
      Y1              =   3720
      Y2              =   4080
   End
   Begin VB.Line airtoptubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   5640
      X2              =   4200
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line airtoptubing 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   7
      Index           =   2
      Visible         =   0   'False
      X1              =   4320
      X2              =   4800
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line airbottubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   4200
      X2              =   5640
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line airbottubing 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   7
      Index           =   2
      Visible         =   0   'False
      X1              =   4320
      X2              =   4560
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line iflowtubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   5880
      X2              =   5880
      Y1              =   3000
      Y2              =   3720
   End
   Begin VB.Line iflowtubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   6480
      X2              =   5880
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line uflowtubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   5880
      X2              =   5880
      Y1              =   3720
      Y2              =   2040
   End
   Begin VB.Line liqbottubing 
      BorderColor     =   &H00000000&
      Index           =   1
      Visible         =   0   'False
      X1              =   5640
      X2              =   4440
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line liqbottubing 
      BorderColor     =   &H00000000&
      Index           =   0
      Visible         =   0   'False
      X1              =   4440
      X2              =   4440
      Y1              =   3480
      Y2              =   5760
   End
   Begin VB.Label mandrainlabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drain"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   113
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape mandrainshape 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line topfilltubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   4440
      X2              =   5160
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line tubing 
      Index           =   3
      X1              =   5640
      X2              =   5640
      Y1              =   3960
      Y2              =   4080
   End
   Begin VB.Line liqtoptubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   5640
      X2              =   5640
      Y1              =   3960
      Y2              =   3480
   End
   Begin VB.Line liqtoptubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   4440
      X2              =   5640
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line airbottubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   4200
      X2              =   4200
      Y1              =   4440
      Y2              =   5040
   End
   Begin VB.Line airtubing 
      Index           =   1
      X1              =   4200
      X2              =   4200
      Y1              =   4440
      Y2              =   3600
   End
   Begin VB.Line liquidtubing 
      Index           =   4
      Visible         =   0   'False
      X1              =   4440
      X2              =   4440
      Y1              =   3240
      Y2              =   3480
   End
   Begin VB.Label liquidlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Penetrometer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4560
      TabIndex        =   112
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line liquidtubing 
      Index           =   3
      Visible         =   0   'False
      X1              =   4320
      X2              =   2640
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line liquidtubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   2400
      X2              =   2520
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line liquidtubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   2640
      X2              =   2640
      Y1              =   3600
      Y2              =   3000
   End
   Begin VB.Line lpgtubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   2040
      X2              =   2040
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line lpgtubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   2040
      X2              =   2040
      Y1              =   3480
      Y2              =   3600
   End
   Begin VB.Line tubing 
      Index           =   1
      X1              =   1560
      X2              =   2640
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line v1tubing 
      Index           =   2
      X1              =   1200
      X2              =   1560
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line mftubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   2640
      X2              =   2640
      Y1              =   4200
      Y2              =   3600
   End
   Begin VB.Line mftubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   2280
      X2              =   2640
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line v2tubing 
      Index           =   0
      X1              =   1080
      X2              =   2280
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line v1tubing 
      Index           =   0
      X1              =   1080
      X2              =   1200
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line reg53tubing 
      BorderWidth     =   5
      Index           =   7
      Visible         =   0   'False
      X1              =   720
      X2              =   360
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line reg53tubing 
      Index           =   6
      Visible         =   0   'False
      X1              =   840
      X2              =   840
      Y1              =   5040
      Y2              =   4440
   End
   Begin VB.Line reg53tubing 
      Index           =   5
      Visible         =   0   'False
      X1              =   840
      X2              =   240
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line reg53tubing 
      Index           =   4
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line reg53tubing 
      Index           =   3
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   4080
      Y2              =   3960
   End
   Begin VB.Line reg53tubing 
      Index           =   2
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   4560
      Y2              =   4320
   End
   Begin VB.Line reg53tubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   4920
      Y2              =   4800
   End
   Begin VB.Line reg53tubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   5280
      Y2              =   5160
   End
   Begin VB.Line tubing 
      Index           =   0
      X1              =   0
      X2              =   600
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V14"
      Height          =   255
      Index           =   14
      Left            =   6720
      TabIndex        =   110
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   14
      Visible         =   0   'False
      X1              =   6720
      X2              =   6960
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line vcloseline 
      Index           =   14
      Visible         =   0   'False
      X1              =   6840
      X2              =   6840
      Y1              =   1920
      Y2              =   2160
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fill"
      Height          =   375
      Index           =   13
      Left            =   5280
      TabIndex        =   108
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line vopenline 
      Index           =   13
      Visible         =   0   'False
      X1              =   5160
      X2              =   5160
      Y1              =   3120
      Y2              =   2880
   End
   Begin VB.Line vcloseline 
      Index           =   13
      Visible         =   0   'False
      X1              =   5280
      X2              =   5040
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drain"
      Height          =   255
      Index           =   12
      Left            =   6480
      TabIndex        =   106
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line vopenline 
      Index           =   12
      Visible         =   0   'False
      X1              =   6360
      X2              =   6360
      Y1              =   5280
      Y2              =   5040
   End
   Begin VB.Line vcloseline 
      Index           =   12
      Visible         =   0   'False
      X1              =   6240
      X2              =   6480
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V11"
      Height          =   255
      Index           =   11
      Left            =   2160
      TabIndex        =   104
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   11
      Visible         =   0   'False
      X1              =   2040
      X2              =   2040
      Y1              =   3240
      Y2              =   3480
   End
   Begin VB.Line vcloseline 
      Index           =   11
      Visible         =   0   'False
      X1              =   1920
      X2              =   2160
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V10"
      Height          =   255
      Index           =   10
      Left            =   1920
      TabIndex        =   102
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   10
      Visible         =   0   'False
      X1              =   2280
      X2              =   2280
      Y1              =   4080
      Y2              =   3840
   End
   Begin VB.Line vcloseline 
      Index           =   10
      Visible         =   0   'False
      X1              =   2400
      X2              =   2160
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V9"
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   100
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   9
      Visible         =   0   'False
      X1              =   2640
      X2              =   2640
      Y1              =   2880
      Y2              =   3000
   End
   Begin VB.Line vcloseline 
      Index           =   9
      Visible         =   0   'False
      X1              =   2640
      X2              =   2520
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V8"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   98
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   8
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   3720
      Y2              =   3960
   End
   Begin VB.Line vcloseline 
      Index           =   8
      Visible         =   0   'False
      X1              =   120
      X2              =   360
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V7"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   96
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   7
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   4080
      Y2              =   4320
   End
   Begin VB.Line vcloseline 
      Index           =   7
      Visible         =   0   'False
      X1              =   120
      X2              =   360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V6"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   94
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   6
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   4560
      Y2              =   4800
   End
   Begin VB.Line vcloseline 
      Index           =   6
      Visible         =   0   'False
      X1              =   360
      X2              =   120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V5"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   92
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   5
      Visible         =   0   'False
      X1              =   240
      X2              =   240
      Y1              =   5160
      Y2              =   4920
   End
   Begin VB.Line vcloseline 
      Index           =   5
      Visible         =   0   'False
      X1              =   120
      X2              =   360
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V4"
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   90
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   4
      Visible         =   0   'False
      X1              =   3600
      X2              =   3840
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line vcloseline 
      Index           =   4
      Visible         =   0   'False
      X1              =   3720
      X2              =   3720
      Y1              =   3480
      Y2              =   3720
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vent"
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   88
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line vopenline 
      Index           =   3
      Visible         =   0   'False
      X1              =   3840
      X2              =   4080
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line vcloseline 
      Index           =   3
      Visible         =   0   'False
      X1              =   3960
      X2              =   3960
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Label ValveLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Drain"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   86
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line vopenline 
      Index           =   2
      Visible         =   0   'False
      X1              =   5640
      X2              =   5640
      Y1              =   5640
      Y2              =   5400
   End
   Begin VB.Line vcloseline 
      Index           =   2
      Visible         =   0   'False
      X1              =   5760
      X2              =   5520
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V1"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   84
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vopenline 
      Index           =   1
      Visible         =   0   'False
      X1              =   1200
      X2              =   1200
      Y1              =   5160
      Y2              =   4920
   End
   Begin VB.Line vcloseline 
      Index           =   1
      Visible         =   0   'False
      X1              =   1320
      X2              =   1080
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label ValveLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fill"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   82
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line vcloseline 
      Index           =   0
      Visible         =   0   'False
      X1              =   5280
      X2              =   5280
      Y1              =   5160
      Y2              =   5400
   End
   Begin VB.Label iflowlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Integrity Flow"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      TabIndex        =   81
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label upglabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Microflow Pressure"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   80
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label cpglabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Compression Pressure"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7440
      TabIndex        =   79
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lpglabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "P2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1860
      TabIndex        =   78
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label hpglabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "P1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   77
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label hflowlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "High Flow 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      TabIndex        =   76
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label mflowlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "High Flow 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3120
      TabIndex        =   75
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lflowlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Low Flow"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1680
      TabIndex        =   74
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lfclabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   72
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lfclabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LFlow Jump: 20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   68
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label comstatlabel 
      Height          =   255
      Left            =   7440
      TabIndex        =   66
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Shape xhhbox 
      Height          =   1005
      Left            =   7440
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label xhhtitle 
      Caption         =   "Ext. Hyd. Head"
      Height          =   255
      Left            =   7440
      TabIndex        =   50
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label readatlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   255
      Left            =   3960
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label creg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "C. Reg.      Counts"
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
      Height          =   255
      Left            =   2640
      TabIndex        =   38
      Top             =   1350
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Line cregline 
      Visible         =   0   'False
      X1              =   6480
      X2              =   7440
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Auxreading 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Auxiliary Input:"
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
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   1890
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label pistonclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   6120
      TabIndex        =   36
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape piston 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   6240
      Top             =   3720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape piston 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   6120
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cregclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label cregclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   34
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label creglabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Compression Regulator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      TabIndex        =   33
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label templabel 
      Caption         =   "Target Temp"
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
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape hflowshape 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2160
      Top             =   4320
      Width           =   495
   End
   Begin VB.Shape mflowshape 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2520
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape v2shape 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   1800
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label regclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label regclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label helpclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   4
      Left            =   7440
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.Label helpclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Index           =   3
      Left            =   6240
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label v2clik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   15
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label v2clik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   14
      Top             =   4800
      Width           =   315
   End
   Begin VB.Label helplabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Click on item for help in use"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   0
      Width           =   2895
   End
   Begin VB.Line helpline 
      Index           =   1
      X1              =   5160
      X2              =   8640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line helpline 
      Index           =   0
      X1              =   5160
      X2              =   5160
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Shape helpshape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape helpshape 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   7080
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape helpshape 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   7440
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape helpshape 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   7080
      Shape           =   2  'Oval
      Top             =   240
      Width           =   930
   End
   Begin VB.Shape helpshape 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   5760
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape helpshape 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Index           =   2
      Left            =   5760
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape helpshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape helpshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   5400
      Top             =   360
      Width           =   135
   End
   Begin VB.Label reglabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Regulator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Shape upgshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape upgshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   6360
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape iflowshape 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   6000
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape iflowshape 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Index           =   0
      Left            =   6000
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape liqpermshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   4320
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape lpgshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape lpgshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   1980
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape hflowshape 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   375
      Index           =   0
      Left            =   2400
      Top             =   4320
      Width           =   255
   End
   Begin VB.Shape mflowshape 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   375
      Index           =   0
      Left            =   2760
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape lflowshape 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1080
      Top             =   4200
      Width           =   495
   End
   Begin VB.Shape lflowshape 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   375
      Index           =   0
      Left            =   1320
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape fixedneedleshape 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   0
      Left            =   1200
      Shape           =   2  'Oval
      Top             =   4680
      Width           =   450
   End
   Begin VB.Shape regshape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   600
      Shape           =   3  'Circle
      Top             =   5040
      Width           =   495
   End
   Begin VB.Shape hpgshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   375
   End
   Begin VB.Shape scshape 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   1
      Left            =   5040
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Shape scshape 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   4920
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Penetro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Penetrometer:"
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
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1620
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Regulator 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Regulator      counts"
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
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1350
      Width           =   5055
   End
   Begin VB.Label Valve_Pos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Valve 2"
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
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label Press_Read 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   810
      Width           =   5055
   End
   Begin VB.Label highflow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "High Flow Rate"
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
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   5055
   End
   Begin VB.Label lowflow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Low Flow Rate"
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
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   270
      Width           =   5055
   End
   Begin VB.Label TLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "00:00:00"
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
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Tag             =   "go"
      Top             =   0
      Width           =   1935
   End
   Begin VB.Shape hpgshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   1560
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape cregshape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape cpgshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   7080
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape cpgshape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   7200
      Top             =   3000
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lflowclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label mflowclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label hflowclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label hpgclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   615
      Left            =   1440
      TabIndex        =   19
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lpgclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   615
      Left            =   1860
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cpgclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   615
      Left            =   7080
      TabIndex        =   37
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label upgclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   615
      Left            =   6240
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label iflowclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   6000
      TabIndex        =   23
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line vopenline 
      Index           =   0
      Visible         =   0   'False
      X1              =   5160
      X2              =   5400
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   83
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   14
      Left            =   6720
      TabIndex        =   111
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   97
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   91
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   89
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   14
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   13
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   12
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   11
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   10
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   120
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   99
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   95
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   93
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   85
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape fixedneedleshape 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   1
      Left            =   1320
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label helpclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Index           =   1
      Left            =   5760
      TabIndex        =   12
      Top             =   240
      Width           =   375
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   10
      Left            =   2160
      TabIndex        =   103
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line v2tubing 
      Index           =   1
      X1              =   2280
      X2              =   2280
      Y1              =   3600
      Y2              =   5280
   End
   Begin VB.Line v1tubing 
      Index           =   1
      X1              =   1200
      X2              =   1200
      Y1              =   3600
      Y2              =   5280
   End
   Begin VB.Shape v2shape 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   1200
      Top             =   4800
      Width           =   855
   End
   Begin VB.Shape v2shape 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   11
      Left            =   1920
      TabIndex        =   105
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   101
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   13
      Left            =   5040
      TabIndex        =   109
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   87
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   12
      Left            =   6240
      TabIndex        =   107
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line auxbottubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   4680
      X2              =   4680
      Y1              =   5040
      Y2              =   3720
   End
   Begin VB.Line liqtoptubing 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   7
      Index           =   2
      Visible         =   0   'False
      X1              =   5640
      X2              =   5640
      Y1              =   3840
      Y2              =   3600
   End
   Begin VB.Line auxbottubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   4680
      X2              =   5880
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label helpclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   615
      Index           =   0
      Left            =   5280
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.Label helpclik 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   114
      Top             =   360
      Width           =   255
   End
   Begin VB.Line topfilltubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   5160
      X2              =   5160
      Y1              =   2760
      Y2              =   3360
   End
   Begin VB.Line uflowtubing 
      Index           =   0
      Visible         =   0   'False
      X1              =   7080
      X2              =   5880
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line drain12tubing 
      Visible         =   0   'False
      X1              =   6360
      X2              =   6360
      Y1              =   4920
      Y2              =   5640
   End
   Begin VB.Line tubing 
      Index           =   2
      X1              =   5640
      X2              =   5640
      Y1              =   4920
      Y2              =   6240
   End
   Begin VB.Line botfilltubing 
      Visible         =   0   'False
      X1              =   5640
      X2              =   4920
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line airtoptubing 
      Index           =   1
      Visible         =   0   'False
      X1              =   3720
      X2              =   4200
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line airtubing 
      Index           =   0
      X1              =   2640
      X2              =   4200
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   124
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   15
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line liquidtubing 
      Index           =   2
      Visible         =   0   'False
      X1              =   2640
      X2              =   2640
      Y1              =   2280
      Y2              =   2880
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   9
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   16
      Left            =   5640
      TabIndex        =   126
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   16
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line recircline1 
      Visible         =   0   'False
      X1              =   5760
      X2              =   5760
      Y1              =   1560
      Y2              =   4080
   End
   Begin VB.Label ValveClik 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   17
      Left            =   5760
      TabIndex        =   129
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line vcloseline 
      Index           =   17
      Visible         =   0   'False
      X1              =   5760
      X2              =   6000
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line vopenline 
      Index           =   17
      Visible         =   0   'False
      X1              =   5880
      X2              =   5880
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Shape ValveFill 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   17
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu Mainmenu 
      Caption         =   "E&xit"
      Index           =   1
   End
   Begin VB.Menu Mainmenu 
      Caption         =   "&Timer"
      Index           =   2
      Begin VB.Menu TMenu 
         Caption         =   "&Stop"
         Index           =   1
      End
      Begin VB.Menu TMenu 
         Caption         =   "&Go"
         Index           =   2
      End
      Begin VB.Menu TMenu 
         Caption         =   "&Reset"
         Index           =   3
      End
   End
   Begin VB.Menu Mainmenu 
      Caption         =   "&Calibrate"
      Index           =   3
   End
   Begin VB.Menu Mainmenu 
      Caption         =   "LV Man Ctrl"
      Index           =   4
   End
   Begin VB.Menu Mainmenu 
      Caption         =   "&Help"
      Index           =   5
   End
End
Attribute VB_Name = "ManualControl_Mettler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim XTimer As Single
Dim ts$(32)                 ' Text strings for this form

Private Sub do_command(Index As Integer)

If command_issued%(Index) = 0 Then
    If (Index < 30 Or Index > 33) And (Index < 49 Or Index > 51) Then
        Rem regulator increments can be backed up
        Rem all other commands can't be executed while another of the same command is pending
        command_issued%(Index) = 1
    End If
    pending$ = pending$ + Chr$(Index)
End If

End Sub

Private Sub chamberiso_Click(Index As Integer)
    chamberiso(Index).Enabled = False
    pending$ = pending$ + Chr$(20 + Index)
End Sub

Private Sub Command1_Click()
pdebug(0).Caption = ""
pdebug(1).Caption = ""
End Sub

Private Sub Command2_Click()
sample_zero_point = last_penetrometer_reading
WPPS "capstuff", "sample_zero_point", Str$(sample_zero_point), CSFile$
End Sub

Private Sub Command3_Click()
bottom_fill_point = last_penetrometer_reading
WPPS "capstuff", "bottom_fill_point", Str$(bottom_fill_point), CSFile$
End Sub

Private Sub Command4_Click()

    Autocal.Show 1
    
End Sub

Private Sub Command5_Click()
do_command 64
End Sub

Private Sub cpgclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        future_CPress% = 0
    Else
        future_CPress% = 1
    End If
    auto_index% = 9
    manual_aux_click = 3 ' signify compression pressure was last clicked on

End Sub

Private Sub cregclik_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        do_command 46 ' clear regulator - same as rewind
        cregrew.Enabled = False
        cregff.Enabled = False
    ElseIf Shift = 1 Then
        do_command 51 + Index ' inc or dec by 10
    Else
        do_command 49 + Index ' inc or dec by 1
    End If
    
End Sub

Private Sub cregff_Click()

    cregrew.Enabled = False
    cregff.Enabled = False
    do_command 48

End Sub

Private Sub cregrew_Click()

    cregrew.Enabled = False
    cregff.Enabled = False
    do_command 46

End Sub

Private Sub cregscroll_Change()
    creglabel.Caption = ts$(1) + ":" + Str$(cregscroll.Value)   ' "Comp. Reg Jump"
End Sub

Private Sub cregstop_Click()
    cregstop.Enabled = False
    do_command 47
End Sub

Private Sub hflowclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' this handles the high flow meter, which could be high flow 2 if there is medium flow meter

    If Not xhflow Then
        If Button = 2 Then
            future_hflow% = 1
        Else
            future_hflow% = 0
        End If
    Else
        If Button = 2 Then
            future_hflow% = 3
        Else
            future_hflow% = 2
        End If
        ' 6.71.20 begin
        If Shift = 1 Then
            suspend_v10 = True
        Else
            suspend_v10 = False
            ' show valve 10 open
            show_valve_open 10
        End If
        ' 6.71.20 end
    End If
    auto_index% = 2

End Sub

Private Sub hpgclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not ExtraPG Then
        If Button = 2 Then
            future_pres% = 1
        Else
            future_pres% = 0
        End If
    Else
        If Button = 2 Then
            future_pres% = 1
            If Shift <> 1 Then
                Rem force valve 11 closed
                do_command 15
            End If
        Else
            future_pres% = 0
            If Shift <> 1 Then
                Rem force valve 11 closed
                do_command 15
            End If
        End If
    End If
    auto_index% = 3
    
End Sub

Private Sub iflowclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' valve 20 is now controlled independently
' we used to close valve 20 on an airtop machine
' and if valve 20 was already closed we would switch to the low flow meter
' we may want to draw a line from the integrity flow meter, maybe not

    future_lflow% = 1 ': lowflow.Caption = "Integrity Flow Rate: "
    auto_index% = 1
    
End Sub

Private Sub lfcscroll_Change()
    lfclabel(0).Caption = ts$(2) + ":" + Str$(lfcscroll.Value)       ' "LFlow Jump"
End Sub

Private Sub lfctrl_Click(Index As Integer)
' 61 is up (index 0)
' 62 is down (index 1)
' 63 is zero (index 2)

    do_command 61 + Index

End Sub

Private Sub lflowclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    future_lflow% = 1
'    If integrity Then
        ' maybe we will show the integrity line, maybe not
'    End If
Else
'    If integrity Then
        ' maybe we will erase the integrity line, maybe not
'    End If
    future_lflow% = 0
End If
auto_index% = 1

End Sub

Private Sub lpgclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        future_pres% = 2
        If Shift <> 1 Then
            do_command 16
        End If
    Else
        future_pres% = 3
        If Shift <> 1 Then
            do_command 16
        End If
    End If
    auto_index% = 3

End Sub

Private Sub mflowclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' there must be an xhflow meter installed for this to be visible

    If Button = 2 Then
        future_hflow% = 1
    Else
        future_hflow% = 0
    End If
    ' 6.71.20 begin
    If Shift = 1 Then
        suspend_v10 = True
    Else
        suspend_v10 = False
        ' show valve 10 closed
        show_valve_closed 10
    End If
    ' 6.71.20 end
    auto_index% = 2

End Sub

Private Sub new_debug_Click()
' do debugging routine
    
    do_command 36

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Dim Key$
    
    Rem any key pressed while in auto_mode will cause valve to stop
    If auto_mode% = 1 And command_issued%(41) = 0 Then
        command_issued%(41) = 1
        Exit Sub
    End If
    Key$ = UCase$(Chr$(KeyAscii))
    Select Case Key$
        Case "A"
            Cancel_Aborted = True
        Case "T"
            'Valve_Test
            do_command 37
        Case "R"
            'ReZeroAll
            do_command 38
        Case "I"
            If Chr$(KeyAscii) = "I" Then
                do_command 32
            Else
                do_command 30
            End If
        Case "D"
            If Chr$(KeyAscii) = "D" Then
                do_command 33
            Else
                do_command 31
            End If
        Case "O"
            do_command 39
        Case "C"
            do_command 40
        'Case "S"
        '    do_command 41
    End Select
    
End Sub

Private Sub Form_Load()

    Dim i%
    Dim ret$, a$
    
    LoadTextStrings
    If low_flow_controller Then
        lfclabel(0).Visible = True
        lfclabel(1).Visible = True
        lfcscroll.Visible = True
        lfctrl(0).Visible = True
        lfctrl(1).Visible = True
        lfctrl(2).Visible = True
        lfclabel(1).Caption = Format$(lfcpos)
        fixedneedleshape(0).Visible = False
        fixedneedleshape(1).Visible = False
    End If
    readatlabel.Visible = readatenabled
    readatcheck.Visible = readatenabled
    new_debug.Visible = debug_button_enable
    mainmenu(4).Visible = lvperm_enable
    regselframe.Visible = dualregulator
    If dualregulator Then
        regsel(Vpos(17)).Value = True
    End If
    pen2frame.Visible = Second_Penetrometer
    If Second_Penetrometer Then
        penselect(penetrometer_select - 1).Value = True
    End If
    If externalhydrohead Then
        xhhbox.Visible = True
        xhhtitle.Visible = True
        xhhvalve(0).Visible = True
        xhhvalve(1).Visible = True
        xhhvalve(2).Visible = True
        Move_Valve 11, "C"
        Move_Valve 12, "C"
        Move_Valve 13, "C"
    End If
    Rem store normal closed left position of v2 plunger
    v2_plunger_left% = v2shape(0).Left
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    pending$ = ""
    want_to_quit_manual_control = False
    For i% = 1 To 64
        command_issued%(i%) = 0
    Next i%
    TLabel.ForeColor = RGB(255, 255, 255)
    TLabel.BackColor = 0
    'RemoveSysMenu Me
    XTimer = Timer
    TLabel.Tag = "stop"
    Timer1.Enabled = True
    ' these are for the text display, not the picture
    If H2OPERM Or DiffPG Then Penetro.Visible = True
    If H2OPERM And DiffPG Then Auxreading.Visible = True
    'If H2OPERM Then
    '    Command2.Visible = True
    '    Command3.Visible = True
    'Else
        Command2.Visible = False
        Command3.Visible = False
    'End If
    If autocompress Then Auxreading.Visible = True
    If auxin And Not Second_Penetrometer Then Auxreading.Visible = True

    If chambers > 2 Then
        For i% = 0 To chambers - 1
            chamberiso(i%).Visible = True
        Next i%
    End If
    If lflow% = 0 Then
        lowflow.Caption = ts$(4) + " " + ts$(3) + "   (" + ts$(5) + "): " ' "Low Flow Rate (High)"
    Else
        lowflow.Caption = ts$(4) + " " + ts$(3) + "   (" + ts$(4) + "): " ' "Low Flow Rate (Low )"
    End If
    
    If Not xhflow Then
        If HFLOW% = 0 Then
            highflow.Caption = ts$(5) + " " + ts$(3) + "   (" + ts$(5) + "): "    ' "High Flow Rate (High)" + ": "
        Else
            highflow.Caption = ts$(5) + " " + ts$(3) + "   (" + ts$(4) + "): "    ' "High Flow Rate (Low )" + ": "
        End If
    Else
        If HFLOW% = 0 Then
            highflow.Caption = ts$(5) + " " + ts$(3) + " (" + ts$(5) + "-1): "    ' "High Flow Rate (Hi-1): "
        ElseIf HFLOW% = 1 Then
            highflow.Caption = ts$(5) + " " + ts$(3) + " (" + ts$(4) + "-1): "    ' "High Flow Rate (Lo-1): "
        ElseIf HFLOW% = 2 Then
            highflow.Caption = ts$(5) + " " + ts$(3) + " (" + ts$(5) + "-2): "    ' "High Flow Rate (Hi-2): "
        Else
            highflow.Caption = ts$(5) + " " + ts$(3) + " (" + ts$(4) + "-2): "    ' "High Flow Rate (Lo-2): "
        End If
    End If

    If HFLOW% > 1 Then
        Move_Valve 9, "O"
    Else
        Move_Valve 9, "C"
    End If
    
    If Not ExtraPG Then
        If Pres% = 0 Then
            Press_Read.Caption = ts$(6) + " (" + ts$(5) + "): " ' "Pressure (High): "
        Else
            Press_Read.Caption = ts$(6) + " (" + ts$(4) + "): " '"Pressure (Low ): "
        End If
    Else
        If Pres% = 0 Then Press_Read.Caption = ts$(5) + " " + ts$(6) + " (" + ts$(5) + "): "    ' "High Pressure (High): "
        If Pres% = 1 Then Press_Read.Caption = ts$(5) + " " + ts$(6) + " (" + ts$(4) + "): "    ' High pressure low
        If Pres% = 2 Then Press_Read.Caption = ts$(4) + " " + ts$(6) + " (" + ts$(5) + "): "    ' Low pressure high
        If Pres% = 3 Then Press_Read.Caption = ts$(4) + " " + ts$(6) + " (" + ts$(4) + "): "    ' Low pressure low
        If Pres% < 2 Then
            Move_Valve 10, "C"
        End If
    End If
    
    If Not newreg Then
        Regulator.Caption = ts$(7) + ": " + Xformat$(REGPOS, "####0 " + ts$(32))    ' "Regulator"/"\C\o\u\n\t\s"
    Else
        update_regulator_display
        regrew.Visible = True
        regff.Visible = True
        regstop.Visible = True
    End If
    
    If autocompress Then
        creglabel.Visible = True
        cpglabel.Visible = True
        pistonlabel.Visible = True
        If ip_creg_enable Then
            creglabel.Caption = ts$(1) + ": 20"     ' "Comp. Reg Jump"
            cregscroll.Visible = True
        Else
            cregrew.Visible = True
            cregff.Visible = True
            cregstop.Visible = True
        End If
        cregshape.Visible = True
        cregclik(0).Visible = True
        cregclik(1).Visible = True
        pistonclik.Visible = True
        piston(0).Visible = True
        cregline.Visible = True
        cpgclik.Visible = True
        cpgshape(0).Visible = True
        cpgshape(1).Visible = True
        creg.Visible = True
    End If
    If autopiston Then
        pistonlabel.Visible = True
        pistonclik.Visible = True
        piston(0).Visible = True
    End If
    
    If H2OPERM Then Penetro.Caption = ts$(8) + ": "         '"Penetrometer"
    If DiffPG Then
        If H2OPERM Then
            Auxreading.Caption = ts$(9) + " (" + ts$(5) + "): " ' "Diff. Press. (High): "
        Else
            Penetro.Caption = ts$(9) + " (" + ts$(5) + "): "    ' "Diff. Press. (High): "
        End If
    End If
    
    If autocompress Then
        Auxreading.Caption = ts$(10) + " (" + ts$(5) + "): " ' "Comp. Press. (High: "
    End If

    If Not RUNNING Then
        Move_Valve 0, "C"
        Move_Valve 2, "C"

        If H2OPERM Then
            Move_Valve 3, "C"
            Move_Valve 8, "C"
            If Drain12 Then
                Move_Valve 11, "C"
            End If
        End If
        If chambers > 2 Then
            ' close all valves
            For i% = 1 To chambers
                Move_Valve -i%, "C"
            Next i%
        End If
        If autocompress And recirculation = False Then
            If safetyup And Vpos(15) = 1 Then
                If safetyupdoor Then
                    ' do auto door switch thing instead of key press thing
                    check_safety_door False
                Else
                    safetykeypress.mainlabel.Caption = ts$(11)      ' "Piston About To Rise"
                    safetykeypress.cancelbutton.Visible = False
                    safetykeypress.Show 1
                End If
            End If
            Move_Valve 14, "C"
        End If
    Else
        ' we are running a test
        If chambers > 2 Then
            ' set the chambers check box to reflect
            ' which unit is current
            For i% = 1 To chambers
                If Vpos(1 - i%) = 1 Then
                    chamberiso(i% - 1).Value = 1
                    ' when we set the value, it will disable
                    ' the button, so we need to re-enable it
                    chamberiso(i% - 1).Enabled = True
                End If
            Next i%
            ' if we set any value, it would add a command
            ' to the pending string, so clear it just
            ' in case
            pending$ = ""
        End If
    End If
    If Vpos(15) = 1 Then piston(1).Visible = True
    RUNNING = False

    'Form_Paint
    If DiffPG Then
        upgclik.Visible = True
        uflowtubing(0).Visible = True
        uflowtubing(1).Visible = True
        upglabel.Visible = True
        upgshape(0).Visible = True
        upgshape(1).Visible = True
        make_valve_visible 14
        If Vpos(14) <> 0 Then
            show_valve_open 14
        End If
        ' deal with frazierpressuregauge
    End If
    If H2OPERM Then
        If Not liqpermonly Then
            make_valve_visible 4
            ' valve 4 is backwards because it used to be a 3-way valve
            If Vpos(4) <> 1 Then
                show_valve_open 4
            End If
        End If
        make_valve_visible 9
        If Vpos(9) <> 0 Then
            show_valve_open 9
        End If
        liqpermshape.Visible = True
        liquidlabel.Visible = True
        liquidtubing(0).Visible = True
        liquidtubing(1).Visible = True
        liquidtubing(2).Visible = True
        liquidtubing(3).Visible = True
        liquidtubing(4).Visible = True
        If Drain12 Then
            drain12tubing.Visible = True
            make_valve_visible 12
            If Vpos(12) <> 0 Then
                show_valve_open 12
            End If
        End If
        If Second_Penetrometer And penetrometer_select = 2 Then
            set_liqtubing (P2PEN20500 >= 0)
        Else
            set_liqtubing (PEN20500 >= 0)
        End If
    End If
    ' stuff for valved regulator
    If (Not newreg) And (Not ip_reg_enable) Then
        For i% = 5 To 8
            make_valve_visible i%
        Next i%
        For i% = 0 To 7
            reg53tubing(i%).Visible = True
        Next i%
    End If
    If ip_reg_enable Then
        reglabel.Caption = ts$(12) + ": 20"     ' "Reg Jump"
        RegScroll.Visible = True
    End If
    If ExtraPG Then
        make_valve_visible 11
        If Vpos(11) <> 0 Then
            show_valve_open 11
        End If
        lpgclik.Visible = True
        lpglabel.Visible = True
        lpgshape(0).Visible = True
        lpgshape(1).Visible = True
        lpgtubing(0).Visible = True
        lpgtubing(1).Visible = True
    End If
    If xhflow Then
        mflowclik.Visible = True
        mflowlabel.Visible = True
        mflowshape(0).Visible = True
        mflowshape(1).Visible = True
        mftubing(0).Visible = True
        mftubing(1).Visible = True
        make_valve_visible 10
        If Vpos(10) <> 0 Then
            show_valve_open 10
        End If
    Else
        hflowlabel.Caption = ts$(13)            ' "High Flow"
        ' it isn't hflow 2 if there is only one of them
    End If
    If itester Or BPTester Or liqpermonly Then
        ' turn off high flow meter
        hflowclik.Visible = False
        hflowlabel.Visible = False
        hflowshape(0).Visible = False
        hflowshape(1).Visible = False
    End If
    ' we still need to handle the liquid perm only case
    If nov2 Then
        Valve_Pos(2).Visible = False
        v2label.Visible = False
        v2clik(0).Visible = False
        v2clik(1).Visible = False
        v2shape(0).Visible = False
        v2shape(1).Visible = False
        v2shape(2).Visible = False
        If Not liqpermonly Then
            v2tubing(0).Visible = False
            v2tubing(1).Visible = False
        End If
    End If
    ' version 6 autofill is from below and uses the valve near the drain valve
    ' (actually it uses a double acting valve for both drain and fill)
    If Auto_fill And Version < 7 Then
        make_valve_visible 0
        botfilltubing.Visible = True
    End If
    If Auto_fill And Version >= 7 Then
        ' version 7 autofill is a top fill valve
        make_valve_visible 13
        topfilltubing(0).Visible = True
        topfilltubing(1).Visible = True
    End If
    If integrity Then
        Rem integrity flow meter added
        iflowclik.Visible = True
        iflowlabel.Visible = True
        iflowshape(0).Visible = True
        iflowshape(1).Visible = True
        iflowtubing(0).Visible = True
        iflowtubing(1).Visible = True
    End If
    If integrity Or DiffPG Then
        If AirTop Then
            auxbottubing(0).Visible = True
            auxbottubing(1).Visible = True
            auxbottubing(2).Visible = True
        Else
            auxtoptubing.Visible = True
        End If
    End If
    If GasPerm Then ' gasperm is true if there is no low flow meter
        ' hide low flow meter, v1 tubing, don't show valve 1
        lflowclik.Visible = False
        lflowlabel.Visible = False
        lflowshape(0).Visible = False
        lflowshape(1).Visible = False
        fixedneedleshape(0).Visible = False
        fixedneedleshape(1).Visible = False
        v1tubing(0).Visible = False
        v1tubing(1).Visible = False
        v1tubing(2).Visible = False
    Else
        make_valve_visible 1
        If Vpos(1) <> 0 Then
            show_valve_open 1
        End If
    End If
    If AirTop Or liqpermonly Or PEN20500 < 0 Then
        ' both have an upper valve 3 for venting
        make_valve_visible 3
        If Vpos(3) <> 0 Then
            show_valve_open 3
        End If
        ' one airtop tube is needed for liqpermonly upper vent
        airtoptubing(1).Visible = True
    End If
    If AirTop Then
        airtoptubing(0).Visible = True
        airtoptubing(2).Visible = True
        If v20_exists Then
            make_valve_visible 2
            If Vpos(20) <> 0 Then
                show_valve_open 2
            End If
        End If
    ElseIf liqpermonly Then
        ' valve picture 2 is used for valve 12 on liqpermonly
        make_valve_visible 2
        ValveLabel(2).Caption = ts$(14)     ' "Isolation"
        If Vpos(12) <> 0 Then
            show_valve_open 2
        End If
        ' show manual drain valve
        mandrainlabel.Visible = True
        mandrainshape.Visible = True
    Else
        airbottubing(0).Visible = True
        airbottubing(1).Visible = True
        airbottubing(2).Visible = True
        make_valve_visible 2
        If PEN20500 < 0 Then
            ' valve 2 is actually valve 12, not valve 3
            If Vpos(12) <> 0 Then
                show_valve_open 2
            End If
        Else
            If Vpos(3) <> 0 Then
                show_valve_open 2
            End If
        End If
    End If
    If temperature% > 0 Then
        templabel.Visible = True
        tempset.Visible = True
        tempval.Visible = True
        If temperature% > 1 Then
            tempprobe1.Visible = True
            tempprobe2.Visible = True
        End If
    End If
    TitleScrn.MousePointer = 0
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(15) + Str$(unitnumber)    ' "Unit"
    End If
    If recirculation Then
        Me.recircline1.Visible = True
        make_valve_visible 15
        If Vpos(21) <> 0 Then
            show_valve_open 15
        End If
        make_valve_visible 16
        If Vpos(22) <> 0 Then
            show_valve_open 16
        End If
    End If
    If valve_23_exists Then
        make_valve_visible 17
        If Vpos(23) <> 0 Then
            show_valve_open 17
        End If
    End If
    
    If doorlock Then Command5.Visible = True
    
    ' Visibility of autocal button
    ret$ = String$(255, 0)
    GPPS "Capstuff", "CalibrateWindow", "", ret$, 255, CSFile$
    a$ = nulltrim(ret$)
    Command4.Visible = (a$ = "Y")
    
    ' initialize future variables.  -1 means no current request
    future_pres% = -1
    future_lflow% = -1
    future_hflow% = -1
    future_DPress% = -1
    future_CPress% = -1
    
End Sub

Private Sub helpclik_Click(Index As Integer)

Dim msg$

Select Case Index
    Case 0:
        msg$ = ts$(16) + vbCrLf + ts$(17)       ' "Click left Button for high range."/"Click right Button for low range."
        MsgBox msg$, 64, ts$(27)                ' "Pressure Gauge"
    Case 1:
        msg$ = ts$(16) + vbCrLf + ts$(17)       ' "Click left button for high range."/"Click right button for low range."
        MsgBox msg$, 64, ts$(28)                ' "Flow Meter"
    Case 2:
        msg$ = ts$(18) + vbCrLf + ts$(19)       ' "Click on valve to toggle open or closed."/"Not all valves can be controlled this way"
        MsgBox msg$, 64, ts$(29)                ' "Solenoid Valve"
    Case 3:
        msg$ = ts$(20) + " (I)." + vbCrLf + ts$(21) + " (D)." + vbCrLf + ts$(22)      ' "Click left button on upper half to increase")/"Click left button on lower half to decrease"/"Click right Button to Zero."
        If Version = 6 Then
            msg$ = msg$ + vbCrLf + ts$(23)      ' "Shift key multiplies command by 10"
        End If
        MsgBox msg$, 64, ts$(6) + " " + ts$(7)  ' "Pressure Regulator"
    Case 4:
        msg$ = ts$(24) + vbCrLf + ts$(25) + vbCrLf + ts$(26)       ' "Click on left half to pulse open."/"Click on right half to pulse closed."/"Press S, O, or C to STOP, OPEN, or CLOSE fully."
        MsgBox msg$, 64, ts$(30)                ' "Needle Valve 2 ONLY"
    End Select
Exit Sub

End Sub

Private Sub Mainmenu_Click(Index As Integer)

    Dim r As Long
    
    Select Case Index
        Case 1
            Timer1.Enabled = False
            want_to_quit_manual_control = True
        Case 2
        Case 3
            do_command 35
        Case 4
            do_command 60
        Case 5
            r = WinHelp(hwnd, HelpFile$, Help_Context, ByVal 55&)
    End Select
    
End Sub

Private Sub penselect_Click(Index As Integer)

    If penetrometer_select <> Index + 1 Then
        Rem change penetrometers
        do_command 59
    End If
    
End Sub

Private Sub pistonclik_Click()
    do_command 53 ' piston toggle
End Sub

Private Sub regclik_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        do_command 34 ' clear regulator
    ElseIf Shift = 1 Then
        do_command 32 + Index ' inc or dec regulator by 10
    Else
        do_command 30 + Index ' inc or dec regulator by 1
    End If
    
End Sub

Private Sub regff_Click()
    regrew.Enabled = False
    regff.Enabled = False
    do_command 44
End Sub

Private Sub regrew_Click()
    regrew.Enabled = False
    regff.Enabled = False
    do_command 42
End Sub

Private Sub RegScroll_Change()
    reglabel.Caption = ts$(12) + ":" + Str$(RegScroll.Value)    ' "Reg Jump"
End Sub

Private Sub regsel_Click(Index As Integer)
    do_command 54 + Index
End Sub

Private Sub regstop_Click()
    regstop.Enabled = False
    do_command 43
End Sub

Private Sub templabel_Click()
' this label acts as the temperature selector if there is only one probe
If temperature% = 1 Then
    manual_aux_click = 1 ' signify temperature was last clicked on
End If
End Sub

Private Sub tempprobe1_Click()
manual_aux_click = 1 ' signify temperature was last clicked on
End Sub

Private Sub tempprobe2_Click()
manual_aux_click = 1 ' signify temperature was last clicked on
End Sub

Private Sub tempset_Click()

    Dim temptarget%

    temptarget% = myVal(tempval.Text) * 10
    If temptarget% < 0 Then temptarget% = 0
    If temptarget% > 9999 Then temptarget% = 9999
    tempval.Text = Str$(temptarget% / 10#)
    pending$ = pending$ + Chr$(45) + Chr$(temptarget% And 127) + Chr$(Int(temptarget% / 128))

End Sub

Private Sub tempset_KeyPress(KeyAscii As Integer)
    Form_KeyPress (KeyAscii)
    KeyAscii = 0
End Sub

Private Sub tempval_KeyPress(KeyAscii As Integer)

    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> Asc(".") And KeyAscii <> 8 Then
        ' don't need to send keys to the form - form has preview turned on
        'Form_KeyPress (KeyAscii)
        KeyAscii = 0
    End If
    
End Sub

Private Sub Timer1_Timer()

    Dim local_time As Single
    Static r1 As Integer
    Static r2 As Integer
    Dim rsum As Long
    
    rsum = CLng(r1) + CLng(r2) + CLng(readings_counter)
    r1 = r2
    r2 = readings_counter
    readings_counter = 0
    comstatlabel.Caption = Format$(numhangs, "#########0") + ":" + Format$(rsum / 3, "####0.##")
    If intest Then Exit Sub
    Me.Caption = ts$(31) + " - " + time$        ' "Manual Control"
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(15) + Str$(unitnumber)    ' "Unit"
    End If

    'If TLabel.Tag = "reset" Then
    '    XTimer = Timer
    '    TLabel.Tag = "go"
    'End If
    If TLabel.Tag = "start" Then
        XTimer = Timer
        TLabel.Tag = "go"
    End If
    If TLabel.Tag = "go" Then
        local_time = Timer
        If local_time + 0.1 < XTimer Then
            XTimer = XTimer - 86400
        End If
        local_time = local_time - XTimer
        TLabel.Caption = Format$((local_time) / 86400, "hh:mm:ss")
    End If

End Sub

Private Sub TMenu_Click(Index As Integer)

    Select Case Index
        Case 1
            TLabel.Tag = "stop"
        Case 2
            If TLabel.Tag = "stop" Then
                TLabel.Tag = "start"
            Else
                TLabel.Tag = "go"
            End If
        Case 3
            XTimer = Timer
            If TLabel.Tag = "stop" Then
                TLabel.Caption = "00:00:00"
            End If
    End Select

End Sub

Private Sub upgclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        future_DPress% = 0
    Else
        future_DPress% = 1
    End If
    auto_index% = 5
    ' maybe draw a line to the microflow gauge
    If H2OPERM Then ' if there is a penetrometer, then the microflow needs to share
                    ' space with other aux readings
        manual_aux_click = 2 ' signify microflow pressure was last clicked on
    End If

End Sub

Private Sub v2clik_Click(Index As Integer)
    do_command 10 + Index
End Sub

Private Sub ValveClik_Click(Index As Integer)

    ' 10 and 11 are not clickable, others map to same commmand
    ' 15 and 16 and 17 map to different commands
    If Index = 10 Or Index = 11 Then
        Exit Sub
    ElseIf Index = 15 Or Index = 16 Or Index = 17 Then
        do_command Index + 2 '(15=17, 16=18, 17=19)
    Else
        do_command Index
    End If

End Sub

Private Sub xhhvalve_Click(Index As Integer)
    xhhvalve(Index).Enabled = False
    pending$ = pending$ + Chr$(56 + Index)
End Sub
Private Sub make_valve_visible(i As Integer)

    ValveClik(i).Visible = True
    ValveFill(i).Visible = True
    ValveLabel(i).Visible = True
    vcloseline(i).Visible = True
    ' leave the open line hidden

End Sub
Sub show_valve_open(i As Integer)

    ' only do this for valves that are actually visible
    If ValveFill(i).Visible = False Then Exit Sub
    vcloseline(i).Visible = False
    vopenline(i).Visible = True
    ' now for color
    If ValveFill(i).FillColor = vbGreen Or ValveFill(i).FillColor = vbRed Then
        ValveFill(i).FillColor = vbGreen
    Else
        ValveFill(i).FillColor = vbYellow
    End If
    
End Sub
Sub show_valve_closed(i As Integer)

    If ValveFill(i).Visible = False Then Exit Sub
    vcloseline(i).Visible = True
    vopenline(i).Visible = False
    ' now for color
    If ValveFill(i).FillColor = vbGreen Or ValveFill(i).FillColor = vbRed Then
        ValveFill(i).FillColor = vbRed
    Else
        ValveFill(i).FillColor = vbCyan
    End If

End Sub
Sub set_liqtubing(b As Boolean)

    ' true = liqtop
    ' false=liqbot
    liqbottubing(0).Visible = Not b
    liqbottubing(1).Visible = Not b
    liqtoptubing(0).Visible = b
    liqtoptubing(1).Visible = b
    liqtoptubing(2).Visible = b

End Sub

Public Sub LoadTextStrings()
' Load text elements for this form from external .ini file
    
    Dim i As Integer
    ' Manual control screen uses extra fonts for some of its labels, represented here:
    Dim mclabel_font As font_info, mc2_font As font_info
    
    ' Load extra font used for labels. Currently no way to set this font
    ' within the program, but that can be dealt with later if necessary.
    mclabel_font.font = gpps2("default", "mclabel_fontname", language$, "Arial")
    mclabel_font.fontsize = Val(gpps2("default", "mclabel_fontsize", language$, "8"))
    mclabel_font.fontbold = (gpps2("default", "mclabel_fontbold", language$, "N") = "Y")
    mc2_font.font = gpps2("default", "mc2_fontname", language$, "MS Serif")
    mc2_font.fontsize = Val(gpps2("default", "mc2_fontsize", language$, "7"))
    mc2_font.fontbold = (gpps2("default", "mc2_fontbold", language$, "N") = "Y")
    
    ' Form elements
    ' Window title is set at the end
    For i = 1 To 5
        mainmenu(i).Caption = gpps2("manual", "mainmenu" + Str$(i), language$, mainmenu(i).Caption)
    Next i
    For i = 1 To 3
        TMenu(i).Caption = gpps2("manual", "tmenu" + Str$(i), language$, TMenu(i).Caption)
    Next i
    Auxreading.Caption = get_thing("manual", "auxreading", language$, Auxreading.Caption, Auxreading, default_font)
    comstatlabel.font = default_font.font   ' size, bold set in design
    set_fontstuff cpgclik, default_font
    set_fontstuff cpglabel, mclabel_font
    cpglabel.Caption = gpps2("manual", "cpglabel", language$, cpglabel.Caption)
    creg.Caption = get_thing("manual", "creg", language$, creg.Caption, creg, default_font)
    creglabel.Caption = get_thing("manual", "creglabel", language$, creglabel.Caption, creglabel, mclabel_font)
    helplabel.Caption = get_thing("manual", "helplabel", language$, helplabel.Caption, helplabel, mclabel_font)
    hflowlabel.Caption = get_thing("manual", "hflowlabel", language$, hflowlabel.Caption, hflowlabel, mclabel_font)
    highflow.Caption = get_thing("manual", "highflow", language$, highflow.Caption, highflow, default_font)
    iflowlabel.Caption = get_thing("manual", "iflow", language$, iflowlabel.Caption, iflowlabel, mclabel_font)
    lfclabel(0).Caption = get_thing("manual", "lfclabel(0)", language$, lfclabel(0).Caption, lfclabel(0), mclabel_font)
    For i = 0 To 2
        lfctrl(i).Caption = gpps2("manual", "lfctrl" + Str$(i), language$, lfctrl(i).Caption)
    Next i
    lflowlabel.Caption = get_thing("manual", "lflowlabel", language$, lflowlabel.Caption, lflowlabel, mclabel_font)
    liquidlabel.Caption = get_thing("manual", "liquidlabel", language$, liquidlabel.Caption, liquidlabel, mclabel_font)
    lowflow.Caption = get_thing("manual", "lowflow", language$, lowflow.Caption, lowflow, default_font)
    mandrainlabel.Caption = get_thing("manual", "mandrain", language$, mandrainlabel.Caption, mandrainlabel, mclabel_font)
    mflowlabel.Caption = get_thing("manual", "mflowlabel", language$, mflowlabel.Caption, mflowlabel, mclabel_font)
    penselect(0).Caption = get_thing("manual", "penselect0", language$, penselect(0).Caption, penselect(0), mc2_font)
    penselect(1).Caption = get_thing("manual", "penselect1", language$, penselect(1).Caption, penselect(1), mc2_font)
    pistonlabel.Caption = get_thing("manual", "pistonlabel", language$, pistonlabel.Caption, pistonlabel, mclabel_font)
    Press_Read.Caption = get_thing("manual", "press_read", language$, Press_Read.Caption, Press_Read, default_font)
    reglabel.Caption = get_thing("manual", "reglabel", language$, reglabel.Caption, reglabel, mclabel_font)
    Regulator.Caption = get_thing("manual", "regulator", language$, Regulator.Caption, Regulator, default_font)
    templabel.Caption = get_thing("manual", "templabel", language$, templabel.Caption, templabel, default_font)
    tempprobe1.Caption = get_thing("manual", "tempprobe1", language$, tempprobe1.Caption, tempprobe1, default_font)
    tempprobe2.Caption = get_thing("manual", "tempprobe2", language$, tempprobe2.Caption, tempprobe2, default_font)
    tempset.Caption = gpps2("manual", "tempset", language$, tempset.Caption)
    set_fontstuff tempval, default_font
    set_fontstuff TLabel, default_font
    upglabel.Caption = get_thing("manual", "upglabel", language$, upglabel.Caption, upglabel, mclabel_font)
    Valve_Pos(2).Caption = get_thing("manual", "valvepos2", language$, Valve_Pos(2).Caption, Valve_Pos(2), default_font)
    ValveLabel(0).Caption = get_thing("manual", "valvelabel0", language$, ValveLabel(0).Caption, ValveLabel(0), mc2_font)
    ValveLabel(2).Caption = get_thing("manual", "valvelabel2", language$, ValveLabel(2).Caption, ValveLabel(2), mc2_font)
    ValveLabel(3).Caption = get_thing("manual", "valvelabel3", language$, ValveLabel(3).Caption, ValveLabel(3), mc2_font)
    ValveLabel(12).Caption = get_thing("manual", "valvelabel12", language$, ValveLabel(12).Caption, ValveLabel(12), mc2_font)
    ValveLabel(13).Caption = get_thing("manual", "valvelabel13", language$, ValveLabel(13).Caption, ValveLabel(13), mc2_font)
    xhhtitle.Caption = get_thing("manual", "xhhtitle", language$, xhhtitle.Caption, xhhtitle, mc2_font)
    For i = 0 To 2
        xhhvalve(i).Caption = get_thing("manual", "xhhvalve" + Str$(i), language$, xhhvalve(i).Caption, xhhvalve(i), mc2_font)
    Next i
    Penetro.Caption = get_thing("manual", "penetro", language$, Penetro.Caption, Penetro, default_font)
    Auxreading.Caption = get_thing("manual", "auxreading", language$, Auxreading.Caption, Auxreading, default_font)
    set_fontname Command1, default_font
    set_fontname Command2, default_font
    set_fontname Command3, default_font
    set_fontname Command4, default_font
    set_fontname Command5, default_font
    Command4.Caption = gpps2("manual", "autocal", language$, Command4.Caption)
    Command1.Caption = gpps2("manual", "command1", language$, Command1.Caption)
    Command2.Caption = gpps2("manual", "command2", language$, Command2.Caption)
    Command3.Caption = gpps2("manual", "command3", language$, Command3.Caption)
    Command5.Caption = gpps2("manual", "command5", language$, Command5.Caption)
    
    ' Other text
    ts$(1) = gpps2("manual", "ts1", language$, "Comp. Reg Jump")
    ts$(2) = gpps2("manual", "ts2", language$, "LFlow Jump")
    ts$(3) = gpps2("manual", "ts3", language$, "Flow Rate")
    ts$(4) = gpps2("manual", "ts4", language$, "Low")
    ts$(5) = gpps2("manual", "ts5", language$, "High")
    ts$(6) = gpps2("manual", "ts6", language$, "Pressure")
    ts$(7) = gpps2("manual", "ts7", language$, "Regulator")
    ts$(8) = gpps2("manual", "ts8", language$, "Penetrometer")
    ts$(9) = gpps2("manual", "ts9", language$, "Diff. Press.")
    ts$(10) = gpps2("manual", "ts10", language$, "Comp. Press.")
    ts$(11) = gpps2("manual", "ts11", language$, "Piston about to rise")
    ts$(12) = gpps2("manual", "ts12", language$, "Reg. Jump")
    ts$(13) = gpps2("manual", "ts13", language$, "High Flow")
    ts$(14) = gpps2("manual", "ts14", language$, "Isolation")
    ts$(15) = gpps2("manual", "ts15", language$, "Unit")
    ts$(16) = gpps2("manual", "ts16", language$, "Click left button for high range.")
    ts$(17) = gpps2("manual", "ts17", language$, "Click right button for low range.")
    ts$(18) = gpps2("manual", "ts18", language$, "Click on valve to toggle open or closed.")
    ts$(19) = gpps2("manual", "ts19", language$, "Not all valves can be controlled this way")
    ts$(20) = gpps2("manual", "ts20", language$, "Click left button on upper half to increase")
    ts$(21) = gpps2("manual", "ts21", language$, "Click left button on lower half to decrease")
    ts$(22) = gpps2("manual", "ts22", language$, "Click right button to zero.")
    ts$(23) = gpps2("manual", "ts23", language$, "Shift key multiplies command by 10")
    ts$(24) = gpps2("manual", "ts24", language$, "Click on left half to pulse open.")
    ts$(25) = gpps2("manual", "ts25", language$, "Click on right half to pulse closed.")
    ts$(26) = gpps2("manual", "ts26", language$, "Press S, O, or C to STOP, OPEN, or CLOSE fully.")
    ts$(27) = gpps2("manual", "ts27", language$, "Pressure Gauge")
    ts$(28) = gpps2("manual", "ts28", language$, "Flow Meter")
    ts$(29) = gpps2("manual", "ts29", language$, "Solenoid Valve")
    ts$(30) = gpps2("manual", "ts30", language$, "Needle Valve 2 ONLY")
    ts$(31) = gpps2("manual", "ts31", language$, "Manual Control")
    ts$(32) = gpps2("manual", "ts32", language$, "\c\o\u\n\t\s")
    
    ManualControl.Caption = get_thing("manual", "window title", language$, ts$(31), ManualControl, default_font)

End Sub
