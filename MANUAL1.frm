VERSION 5.00
Begin VB.Form ManualControl1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Control"
   ClientHeight    =   7245
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14640
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   8.25
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
   ScaleHeight     =   7245
   ScaleMode       =   0  'User
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Manual Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14415
      Begin VB.CommandButton cregff 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13320
         Picture         =   "MANUAL1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cregstop 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12960
         Picture         =   "MANUAL1.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cregrew 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         Picture         =   "MANUAL1.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   13920
         Top             =   120
      End
      Begin VB.Frame WettingFrame 
         Caption         =   "Auto Wetting Pump Controls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   10800
         TabIndex        =   44
         Top             =   4200
         Width           =   2175
         Begin VB.CommandButton Command7 
            Caption         =   "Pump Stop"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   49
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Pump Start"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkWettingValve1 
            Caption         =   "Wetting Valve #1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkWettingValve2 
            Caption         =   "Wetting Valve #2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   1080
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox chkWettingValve3 
            Caption         =   "Wetting Valve #3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   1440
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdShowWettingControls 
         Caption         =   "Show Full Wetting Controls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         TabIndex        =   43
         Top             =   6240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.HScrollBar cregscroll 
         Height          =   255
         LargeChange     =   10
         Left            =   10800
         Max             =   255
         Min             =   1
         TabIndex        =   41
         Top             =   3720
         Value           =   20
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar lfcscroll 
         Height          =   255
         LargeChange     =   10
         Left            =   4080
         Max             =   255
         Min             =   1
         TabIndex        =   35
         Top             =   6120
         Value           =   20
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton lfctrl 
         Caption         =   "LFLOW ZERO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   32
         Top             =   6120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton lfctrl 
         Caption         =   "LFLOW DOWN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   31
         Top             =   5880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton lfctrl 
         Caption         =   "LFLOW UP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   30
         Top             =   5640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar RegScroll 
         Height          =   255
         LargeChange     =   10
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   28
         Top             =   6000
         Value           =   20
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13200
         TabIndex        =   1
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "V7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   7800
         TabIndex        =   87
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line vcloseline 
         Index           =   27
         Visible         =   0   'False
         X1              =   7920
         X2              =   7920
         Y1              =   3600
         Y2              =   3840
      End
      Begin VB.Line vopenline 
         Index           =   27
         Visible         =   0   'False
         X1              =   7800
         X2              =   8040
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label ValveClik 
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
         Height          =   255
         Index           =   27
         Left            =   7800
         TabIndex        =   86
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblReserveTankLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "Reserve Tank Level:"
         Height          =   255
         Left            =   11280
         TabIndex        =   85
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label lblReserveTankLevelValue 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12600
         TabIndex        =   84
         Top             =   1680
         Width           =   855
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
         Left            =   840
         TabIndex        =   10
         Top             =   5160
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
         Left            =   840
         TabIndex        =   11
         Top             =   4920
         Width           =   495
      End
      Begin VB.Shape hflowshape 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   5160
         Top             =   3960
         Width           =   495
      End
      Begin VB.Shape mflowshape 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   5520
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape hflowshape 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   2640
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label hflowclik 
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
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   4320
         Width           =   495
      End
      Begin VB.Shape mflowshape 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   3000
         Top             =   3480
         Visible         =   0   'False
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
         Left            =   3000
         TabIndex        =   17
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape lflowshape 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   1680
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label lfowclik 
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
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label lowFlowRate4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Low Flow Rate 4"
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
         TabIndex        =   83
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label highFlowRate4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "High Flow Rate 4"
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
         TabIndex        =   82
         Top             =   2085
         Width           =   5055
      End
      Begin VB.Label lowFlowRate3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Low Flow Rate 3"
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
         TabIndex        =   81
         Top             =   1200
         Width           =   5055
      End
      Begin VB.Label highFlowRate3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "High Flow Rate 3"
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
         TabIndex        =   80
         Top             =   1485
         Width           =   5055
      End
      Begin VB.Label lowFlowRate2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Low Flow Rate 2"
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
         TabIndex        =   79
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label highFlowRate2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "High Flow Rate 2"
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
         TabIndex        =   78
         Top             =   885
         Width           =   5055
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MV3"
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
         Left            =   6000
         TabIndex        =   77
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MV2"
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
         Left            =   4320
         TabIndex        =   76
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label mv3click 
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
         Left            =   5640
         TabIndex        =   75
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label mv3click 
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
         Left            =   5400
         TabIndex        =   74
         Top             =   4680
         Width           =   255
      End
      Begin VB.Line Line7 
         X1              =   5640
         X2              =   5640
         Y1              =   4680
         Y2              =   4560
      End
      Begin VB.Line Line6 
         X1              =   5640
         X2              =   4920
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line5 
         X1              =   4920
         X2              =   4920
         Y1              =   4680
         Y2              =   4560
      End
      Begin VB.Line Line4 
         X1              =   5640
         X2              =   5640
         Y1              =   5040
         Y2              =   4920
      End
      Begin VB.Line Line3 
         X1              =   4920
         X2              =   4920
         Y1              =   5040
         Y2              =   4920
      End
      Begin VB.Line Line2 
         X1              =   4920
         X2              =   5640
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line1 
         X1              =   5280
         X2              =   5280
         Y1              =   5160
         Y2              =   5040
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
         Left            =   10920
         TabIndex        =   73
         Top             =   3000
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
         Index           =   1
         Left            =   10920
         TabIndex        =   72
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape piston 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   1
         Left            =   9480
         Top             =   3480
         Visible         =   0   'False
         Width           =   135
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
         Index           =   0
         Left            =   9360
         TabIndex        =   71
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape piston 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   9360
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label newValveLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "V10'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   67
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label newValveClik 
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
         Height          =   255
         Left            =   5160
         TabIndex        =   66
         Top             =   3240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label mv2click 
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
         Left            =   4680
         TabIndex        =   65
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label mv2click 
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
         Left            =   4920
         TabIndex        =   64
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label mv1click 
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
         Left            =   2520
         TabIndex        =   63
         Top             =   4800
         Width           =   315
      End
      Begin VB.Label mv1click 
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
         Left            =   2880
         TabIndex        =   62
         Top             =   4800
         Width           =   255
      End
      Begin VB.Line vcloseline 
         Index           =   3
         Visible         =   0   'False
         X1              =   7320
         X2              =   7320
         Y1              =   4320
         Y2              =   4560
      End
      Begin VB.Line vopenline 
         Index           =   3
         Visible         =   0   'False
         X1              =   7200
         X2              =   7440
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label ValveClik 
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
         Height          =   255
         Index           =   3
         Left            =   7200
         TabIndex        =   61
         Top             =   4320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   7200
         Shape           =   3  'Circle
         Top             =   4320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line airtoptubing 
         Index           =   1
         Visible         =   0   'False
         X1              =   7080
         X2              =   7560
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label ValveClik 
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
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   60
         Top             =   4800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line vcloseline 
         Index           =   1
         Visible         =   0   'False
         X1              =   1920
         X2              =   1680
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line vopenline 
         Index           =   1
         Visible         =   0   'False
         X1              =   1800
         X2              =   1800
         Y1              =   5040
         Y2              =   4800
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   4800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line v1tubing 
         Index           =   1
         X1              =   1800
         X2              =   1800
         Y1              =   3000
         Y2              =   5160
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
         Left            =   7800
         TabIndex        =   59
         Top             =   1155
         Visible         =   0   'False
         Width           =   2415
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
         Left            =   5160
         TabIndex        =   58
         Top             =   1155
         Width           =   5055
      End
      Begin VB.Label piston_position_transducer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Thickness:"
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
         Left            =   5160
         TabIndex        =   57
         Top             =   1440
         Visible         =   0   'False
         Width           =   5055
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
         Left            =   5160
         TabIndex        =   56
         Top             =   2040
         Visible         =   0   'False
         Width           =   5055
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
         Left            =   5160
         TabIndex        =   55
         Top             =   1740
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Label Valve_Pos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Motor Valve"
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
         Left            =   5160
         TabIndex        =   54
         Top             =   875
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
         Left            =   5160
         TabIndex        =   53
         Top             =   585
         Width           =   5055
      End
      Begin VB.Label highFlowRate1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "High Flow Rate 1"
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
         TabIndex        =   52
         Top             =   285
         Width           =   5055
      End
      Begin VB.Label lowFlowRate1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Low Flow Rate 1"
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
         TabIndex        =   51
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label comstatlabel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12600
         TabIndex        =   50
         Top             =   6600
         Width           =   1695
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
         Left            =   10800
         TabIndex        =   42
         Top             =   3480
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   10920
         TabIndex        =   40
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
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
         Left            =   9240
         TabIndex        =   39
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape cregshape 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   10920
         Shape           =   3  'Circle
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape cpgshape 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   10440
         Shape           =   3  'Circle
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape cpgshape 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   1
         Left            =   10560
         Top             =   2880
         Visible         =   0   'False
         Width           =   135
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
         Left            =   10440
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line cregline 
         Visible         =   0   'False
         X1              =   9720
         X2              =   10920
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line tubing 
         Index           =   2
         X1              =   9000
         X2              =   9000
         Y1              =   4800
         Y2              =   5760
      End
      Begin VB.Line tubing 
         Index           =   3
         X1              =   9000
         X2              =   9000
         Y1              =   3840
         Y2              =   3720
      End
      Begin VB.Shape scshape 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   735
         Index           =   1
         Left            =   8520
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Shape scshape 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   0
         Left            =   8400
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Line airtoptubing 
         Index           =   0
         Visible         =   0   'False
         X1              =   9000
         X2              =   7560
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label TLabel 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Label lblDemo 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "DEMO ON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   36
         Top             =   6000
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   4080
         TabIndex        =   34
         Top             =   5520
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
         Left            =   4080
         TabIndex        =   33
         Top             =   5880
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   240
         TabIndex        =   29
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label newhflowlabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "High Flow 4"
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
         Left            =   5760
         TabIndex        =   27
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label newmflowlabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "High Flow 3"
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
         Left            =   6120
         TabIndex        =   26
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Line newvcloseline 
         Visible         =   0   'False
         X1              =   5400
         X2              =   5160
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line newvopenline 
         Visible         =   0   'False
         X1              =   5280
         X2              =   5280
         Y1              =   4080
         Y2              =   3840
      End
      Begin VB.Shape newValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   5160
         Shape           =   3  'Circle
         Top             =   3360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label newmflowclik 
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
         Left            =   5520
         TabIndex        =   25
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape mflowshape 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Index           =   3
         Left            =   5760
         Top             =   3240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line newmftubing 
         Index           =   1
         Visible         =   0   'False
         X1              =   5640
         X2              =   5640
         Y1              =   3840
         Y2              =   3000
      End
      Begin VB.Line newmftubing 
         Index           =   0
         Visible         =   0   'False
         X1              =   5280
         X2              =   5640
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Shape hflowshape 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Index           =   3
         Left            =   5400
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label newhflowclik 
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
         Left            =   5160
         TabIndex        =   24
         Top             =   3960
         Width           =   495
      End
      Begin VB.Shape v2shape 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   4800
         Shape           =   2  'Oval
         Top             =   4680
         Width           =   495
      End
      Begin VB.Shape v2shape 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   4800
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label v2label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MV1"
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
         Left            =   3240
         TabIndex        =   23
         Top             =   4800
         Width           =   375
      End
      Begin VB.Line newv2tubing 
         Index           =   1
         X1              =   5280
         X2              =   5280
         Y1              =   3000
         Y2              =   4560
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Vent V3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7080
         TabIndex        =   22
         Top             =   4560
         Visible         =   0   'False
         Width           =   615
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
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   4440
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
         Height          =   375
         Left            =   3600
         TabIndex        =   20
         Top             =   3480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Line airtubing 
         Index           =   1
         X1              =   7560
         X2              =   7560
         Y1              =   4440
         Y2              =   3000
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "V10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   2280
         TabIndex        =   19
         Top             =   3480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "V1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   18
         Top             =   4680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line airtubing 
         Index           =   0
         X1              =   3120
         X2              =   7560
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line mftubing 
         Index           =   1
         Visible         =   0   'False
         X1              =   3120
         X2              =   3120
         Y1              =   4080
         Y2              =   3000
      End
      Begin VB.Line mftubing 
         Index           =   0
         Visible         =   0   'False
         X1              =   2760
         X2              =   3120
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Shape mflowshape 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Index           =   0
         Left            =   3240
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape v2shape 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   2640
         Shape           =   2  'Oval
         Top             =   4800
         Width           =   495
      End
      Begin VB.Shape v2shape 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   2640
         Top             =   4800
         Width           =   375
      End
      Begin VB.Shape lflowshape 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Index           =   0
         Left            =   1920
         Top             =   4080
         Width           =   255
      End
      Begin VB.Shape hflowshape 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   375
         Index           =   0
         Left            =   2880
         Top             =   4320
         Width           =   255
      End
      Begin VB.Line v2tubing 
         Index           =   2
         X1              =   2520
         X2              =   5280
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line vcloseline 
         Index           =   10
         Visible         =   0   'False
         X1              =   2880
         X2              =   2640
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line vopenline 
         Index           =   10
         Visible         =   0   'False
         X1              =   2760
         X2              =   2760
         Y1              =   4080
         Y2              =   3840
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   10
         Left            =   2640
         Shape           =   3  'Circle
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label ValveClik 
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
         Height          =   255
         Index           =   10
         Left            =   2640
         TabIndex        =   14
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line v2tubing 
         Index           =   1
         X1              =   2760
         X2              =   2760
         Y1              =   3000
         Y2              =   5160
      End
      Begin VB.Label hpgLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "P1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   2520
         Width           =   375
      End
      Begin VB.Line tubing 
         Index           =   1
         X1              =   2040
         X2              =   3120
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Shape hpgshape 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   375
      End
      Begin VB.Shape hpgshape 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   1
         Left            =   1920
         Top             =   2640
         Width           =   135
      End
      Begin VB.Line v1tubing 
         Index           =   2
         X1              =   1800
         X2              =   2040
         Y1              =   3000
         Y2              =   3000
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
         Left            =   1920
         TabIndex        =   12
         Top             =   2400
         Width           =   375
      End
      Begin VB.Line v2tubing 
         Index           =   0
         X1              =   1320
         X2              =   2520
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Shape regshape 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   840
         Shape           =   3  'Circle
         Top             =   4920
         Width           =   495
      End
      Begin VB.Line tubing 
         Index           =   0
         X1              =   240
         X2              =   840
         Y1              =   5160
         Y2              =   5160
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
         Left            =   12480
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.Shape helpshape 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   7
         Left            =   13080
         Shape           =   2  'Oval
         Top             =   480
         Width           =   570
      End
      Begin VB.Shape helpshape 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   8
         Left            =   12960
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape helpshape 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   5
         Left            =   11880
         Shape           =   3  'Circle
         Top             =   480
         Width           =   495
      End
      Begin VB.Shape helpshape 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   3
         Left            =   11400
         Top             =   480
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
         Left            =   11400
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape helpshape 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Index           =   0
         Left            =   11040
         Top             =   720
         Width           =   135
      End
      Begin VB.Shape helpshape 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   1
         Left            =   10920
         Shape           =   3  'Circle
         Top             =   480
         Width           =   375
      End
      Begin VB.Label pdebug 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   12120
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label pdebug 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   11040
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Shape helpshape 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   12480
         Shape           =   3  'Circle
         Top             =   480
         Width           =   255
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
         Index           =   3
         Left            =   11880
         TabIndex        =   6
         Top             =   480
         Width           =   495
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
         Left            =   10920
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
      Begin VB.Line helpline 
         Index           =   1
         X1              =   10800
         X2              =   14280
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line helpline 
         Index           =   0
         X1              =   10800
         X2              =   10800
         Y1              =   240
         Y2              =   1200
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
         Height          =   735
         Index           =   0
         Left            =   10800
         TabIndex        =   4
         Top             =   480
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
         Height          =   495
         Index           =   4
         Left            =   12960
         TabIndex        =   3
         Top             =   360
         Width           =   975
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
         Left            =   11400
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape v2shape 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   5400
         Shape           =   2  'Oval
         Top             =   4680
         Width           =   495
      End
      Begin VB.Shape v2shape 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   5400
         Top             =   4680
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   27
         Left            =   7800
         Shape           =   3  'Circle
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "E&xit"
      Index           =   1
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Timer"
      Index           =   2
      Begin VB.Menu Tmenu 
         Caption         =   "Stop"
         Index           =   1
      End
      Begin VB.Menu Tmenu 
         Caption         =   "Go"
         Index           =   2
      End
      Begin VB.Menu Tmenu 
         Caption         =   "Reset"
         Index           =   3
      End
      Begin VB.Menu Tmenu 
         Caption         =   "Pluse V1"
         Index           =   4
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Calibrate"
      Index           =   3
   End
   Begin VB.Menu mainmenu 
      Caption         =   "LV ManCtrl"
      Index           =   4
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Help"
      Index           =   5
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Status"
      Index           =   6
      Begin VB.Menu statusmenu 
         Caption         =   "Off"
         Index           =   0
      End
      Begin VB.Menu statusmenu 
         Caption         =   "Red"
         Index           =   1
      End
      Begin VB.Menu statusmenu 
         Caption         =   "Yellow"
         Index           =   2
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Test Piston"
      Index           =   7
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Calibrate"
      Index           =   8
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Microflow Volume Select"
      Index           =   9
      Visible         =   0   'False
      Begin VB.Menu selvolume 
         Caption         =   "Volume 1"
         Index           =   1
      End
      Begin VB.Menu selvolume 
         Caption         =   "Volume 2"
         Index           =   2
      End
      Begin VB.Menu selvolume 
         Caption         =   "Volume 3"
         Index           =   3
      End
      Begin VB.Menu selvolume 
         Caption         =   "Use All Volume"
         Index           =   4
      End
   End
End
Attribute VB_Name = "ManualControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim XTimer As Single
Dim ts$(36)                 ' Text strings for this form
Dim reservoirLabel As String
Dim dryprobeLabel As String
Dim wetprobeLabel As String
Dim airprobeLabel As String
Dim bubblerprobeLabel As String
Dim cabinetprobeLabel As String
Dim hydroheadprobeLabel As String
Dim mullenprobeLabel As String
Dim pulseV1Enable As Boolean
Dim status_light_setting As Integer

Dim humidityValveOpenLimit As Long
Dim humidityValveCloseLimit As Long
Dim humidityValvePosition As Long

Dim continuousReadChannel As String


Private Sub do_command(Index As Integer)

    Debug.Print "Do_Command: " + str(Index)
    
    If command_issued%(Index) = 0 Then
        If (Index < 30 Or Index > 33) And (Index < 49 Or Index > 51) Then
            Rem regulator increments can be backed up
            Rem all other commands can't be executed while another of the same command is pending
            command_issued%(Index) = 1
        End If
        pending$ = pending$ + Chr$(Index)
    End If

End Sub

Private Sub air_inlet_option1_Click()
    do_command 67
End Sub

Private Sub air_inlet_option2_Click()
    do_command 68
End Sub

Private Sub BubblerMVclik_Click(Index As Integer)
    do_command 97 + Index
End Sub

Private Sub chamberiso_Click(Index As Integer)
    'chamberiso(Index).Enabled = False
    pending$ = pending$ + Chr$(20 + Index)
End Sub

Private Sub Check1_Click()
    do_command 87
End Sub

Private Sub cmdSetHumidityValve_Click()
'
'    Dim temptarget%
'
'    temptarget% = myVal(txtHumidityValve.Text)
'    If temptarget% < 0 Then
'        temptarget% = 0
'    ElseIf temptarget% > 4000 Then
'        temptarget% = 4000
'    End If
'    tempval.Text = Str$(temptarget%)
'    pending$ = pending$ + Chr$(90) + Chr$(temptarget% And 127) + Chr$(Int(temptarget% / 128))
'

End Sub

Private Sub chkControlBubblerMV_Click()

    'If chkControlBubblerMV.value = 1 Then
        'useBubblerMV = True
        'BubblerMVShape(0).Visible = True
        'BubblerMVShape(1).Visible = True
    'Else
        'useBubblerMV = False
        'BubblerMVShape(0).Visible = False
        'BubblerMVShape(1).Visible = False
    'End If

End Sub

Private Sub chkWettingValve1_Click()
    do_command 87
End Sub

Private Sub cmdShowWettingControls_Click()
    frmWettingControls.Show
End Sub

Private Sub Combo1_Click()
Dim s As String
manual_aux_click = 1
's = Combo1.List(Combo1.ListIndex)
'If s = reservoirLabel Then
    'Combo1.Tag = "R"
'ElseIf s = dryprobeLabel Then
    'Combo1.Tag = "D"
'ElseIf s = wetprobeLabel Then
    'Combo1.Tag = "W"
'ElseIf s = airprobeLabel Then
    'Combo1.Tag = "A"
'ElseIf s = bubblerprobeLabel Then
    'Combo1.Tag = "B"
'ElseIf s = cabinetprobeLabel Then
    'Combo1.Tag = "C"
'End If
End Sub

Private Sub Command1_Click()
pdebug(0).Caption = ""
pdebug(1).Caption = ""
End Sub

Private Sub Command2_Click()
sample_zero_point = last_penetrometer_reading
WPPS "capstuff", "sample_zero_point", str$(sample_zero_point), CSFile$
End Sub

Private Sub Command3_Click()
bottom_fill_point = last_penetrometer_reading
WPPS "capstuff", "bottom_fill_point", str$(bottom_fill_point), CSFile$
End Sub

Private Sub Command4_Click()

    Autocal.Show 1
    
End Sub

Private Sub Command5_Click()
do_command 64
End Sub

Private Sub Command6_Click()
    do_command 85
End Sub

Private Sub Command7_Click()
    do_command 86
End Sub

Private Sub Command8_Click()
    do_command 88
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
    creglabel.Caption = ts$(1) + ":" + str$(cregscroll.value)   ' "Comp. Reg Jump"
End Sub

Private Sub cregstop_Click()
    cregstop.Enabled = False
    do_command 47
End Sub

Private Sub hflowclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' this handles the high flow meter, which could be high flow 2 if there is medium flow meter
    motorValveIndex = 0
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
            Debug.Print "future pres = 1"
        Else
            future_pres% = 0
            Debug.Print "future pres = 0"
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

'Private Sub hregclik_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then
'        do_command 90 ' clear regulator
'    ElseIf Shift = 1 Then
'        do_command 93 + Index ' inc or dec regulator by 10
'    Else
'        do_command 91 + Index ' inc or dec regulator by 1
'    End If
'End Sub
'
'Private Sub hregscroll_Change()
'    'changeregby.Caption = "Increase Regulator By: " + Str(hregscroll.value)
'
'End Sub

Private Sub iflowclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' valve 20 is now controlled independently
' we used to close valve 20 on an airtop machine
' and if valve 20 was already closed we would switch to the low flow meter
' we may want to draw a line from the integrity flow meter, maybe not

    future_lflow% = 1 ': lowflow.Caption = "Integrity Flow Rate: "
    auto_index% = 1
    
End Sub

Private Sub lfcscroll_Change()
    lfclabel(0).Caption = ts$(2) + ":" + str$(lfcscroll.value)       ' "LFlow Jump"
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

Private Sub lowflow_Click()
    'allFlowReadingsFrame.Visible = True
End Sub

Private Sub lowFlowRate1_Click()
    'allFlowReadingsFrame.Visible = False
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
Private Sub Auxreading_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Auxreading.Caption = "P3"
    
End Sub
Private Sub mflowclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' there must be an xhflow meter installed for this to be visible
    motorValveIndex = 0
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

Private Sub mv1click_Click(Index As Integer)
    If motorValveIndex <> 0 Then
        motorValveIndex = 0
        do_command 54
    End If
    do_command 10 + Index
End Sub

Private Sub MV1Option_Click()
'    motorValveIndex = 0
'    do_command 54
    'regsel(0).value = True
    'regsel(1).value = False
End Sub

Private Sub mv1click_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        motorValveIndex = 0
        do_command 54
    ElseIf Button = 2 Then
        If motorValveIndex <> 0 Then
            motorValveIndex = 0
            do_command 54
        End If
        do_command 10 + Index
    End If
End Sub

Private Sub mv2click_Click(Index As Integer)
    If motorValveIndex <> 1 Then
        motorValveIndex = 1
        do_command 55
    End If
    do_command 10 + Index
End Sub

Private Sub MV2Option_Click()
    motorValveIndex = 1
    do_command 55
    'regsel(0).value = False
    'regsel(1).value = True
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
        Case "S"
            do_command 41
    End Select
    
End Sub

Private Sub Form_Load()

    Dim i%
    Dim Ret$, a$
    
    LoadTextStrings
    pulseV1Enable = False
    
    useBubblerMV = False
    
    If low_flow_controller Then
        lfclabel(0).Visible = True
        lfclabel(1).Visible = True
        lfcscroll.Visible = True
        lfctrl(0).Visible = True
        lfctrl(1).Visible = True
        lfctrl(2).Visible = True
        lfclabel(1).Caption = Format$(lfcpos)
        'fixedneedleshape(0).Visible = False
        'fixedneedleshape(1).Visible = False
    End If
    If ComLoc = 0 Then
        lblDemo.Visible = True
    Else
        lblDemo.Visible = False
    End If
        
    'readatlabel.Visible = readatenabled
    'readatcheck.Visible = readatenabled
    'new_debug.Visible = debug_button_enable
    mainmenu(4).Visible = lvperm_enable
    mainmenu(6).Visible = status_lights_enable
    'regselframe.Visible = dualregulator
    If dualregulator Then
        'regsel(Vpos(17)).value = True
    End If
    'pen2frame.Visible = Second_Penetrometer
    If Second_Penetrometer Then
        'penselect(penetrometer_select - 1).value = True
    End If
    If externalhydrohead Then
        'xhhbox.Visible = True
        'xhhtitle.Visible = True
        'xhhvalve(0).Visible = True
        'xhhvalve(1).Visible = True
        'xhhvalve(2).Visible = True
        Move_Valve 11, "C"
        Move_Valve 12, "C"
        Move_Valve 13, "C"
    End If
    If air_inlets = 1 Then
        'air_inlet_frame.Visible = False
    Else
        If current_air_inlet = 1 Then
            'air_inlet_option1.value = True
        ElseIf current_air_inlet = 2 Then
            'air_inlet_option2.value = True
        End If
    End If
    
    Rem store normal closed left position of v2 plunger
    v2_plunger_left% = v2shape(0).Left
    'bubblerMV_plunger_left% = BubblerMVShape(0).Left
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    pending$ = ""
    want_to_quit_manual_control = False
    For i% = 1 To UBound(command_issued%)
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
    If H2OPERM Then
        'Command2.Visible = True
        'Command3.Visible = True
    Else
        'Command2.Visible = False
        'Command3.Visible = False
    End If
    If autocompress Then Auxreading.Visible = True
    If auxin And Not Second_Penetrometer Then Auxreading.Visible = True
    If dryChamberTemperature <> 0 Or wetChamberTemperature <> 0 Or reservoirTemperature <> 0 Then
        Auxreading.Visible = True
    End If
    If multiChamberSystem Then
        For i% = 0 To chambers - 1
            'chamberiso(i%).Visible = True
        Next i%
        ' turn off split in chamber
        'Line1.Visible = False
    Else
        ' unless there are two different chambers (one for air, one for liquid)
        'Line1.Visible = (chambers = 2)
    End If
    If lflow% = 0 Then
        'lowflow.Caption = ts$(4) + " " + ts$(3) + "   (" + ts$(5) + "): " ' "Low Flow Rate (High)"
    Else
        'lowflow.Caption = ts$(4) + " " + ts$(3) + "   (" + ts$(4) + "): " ' "Low Flow Rate (Low )"
    End If
    
    If Not xhflow Then
        If HFLOW% = 0 Then
            'highflow.Caption = ts$(5) + " " + ts$(3) + "   (" + ts$(5) + "): "    ' "High Flow Rate (High)" + ": "
        Else
            'highflow.Caption = ts$(5) + " " + ts$(3) + "   (" + ts$(4) + "): "    ' "High Flow Rate (Low )" + ": "
        End If
    Else
        If HFLOW% = 0 Then
            'highflow.Caption = ts$(5) + " " + ts$(3) + " (" + ts$(5) + "-1): "    ' "High Flow Rate (Hi-1): "
        ElseIf HFLOW% = 1 Then
            'highflow.Caption = ts$(5) + " " + ts$(3) + " (" + ts$(4) + "-1): "    ' "High Flow Rate (Lo-1): "
        ElseIf HFLOW% = 2 Then
            'highflow.Caption = ts$(5) + " " + ts$(3) + " (" + ts$(5) + "-2): "    ' "High Flow Rate (Hi-2): "
        Else
            'highflow.Caption = ts$(5) + " " + ts$(3) + " (" + ts$(4) + "-2): "    ' "High Flow Rate (Lo-2): "
        End If
    End If
    
    If hasHumidityControls Then
        'frameHumidity.Visible = True
        'BubblerMVShape(2).BackColor = frameHumidity.BackColor
    Else
        'frameHumidity.Visible = False
    End If

    If hasMultipleMVs Then
        If HFLOW% <= 1 Then
            'Move_Valve 9, "C"
        End If
    Else
        If HFLOW% > 1 Then
            Move_Valve 9, "O"
        Else
            Move_Valve 9, "C"
        End If
    End If
    
    If Not ExtraPG Then
        If Pres% = 0 Then
            Press_Read.Caption = ts$(6) + " (" + ts$(5) + "): " '"Pressure (High): "
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
        'regrew.Visible = True
        'regff.Visible = True
        'regstop.Visible = True
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
        pistonclik(0).Visible = True
        piston(0).Visible = True
        cregline.Visible = True
        cpgclik.Visible = True
        cpgshape(0).Visible = True
        cpgshape(1).Visible = True
        creg.Visible = True
    End If
    If autopiston Or FrazierPiston Then
        pistonlabel.Visible = True
        pistonclik(0).Visible = True
        piston(0).Visible = True
    End If
    If gpps2("Capstuff", "manualMultiChamber", CSFile$, "N") = "Y" Then
            'pistonclik2(1).Visible = True
            'piston(2).Visible = True
    End If
    
'    If DiffPG Then directly follows inserted code
' **********
' BEGIN code inserted by search for Tim Richards Friday 6/4/04
' conditional If H2OPERM Then Penetro.Caption = ts$(8) + ": "         '"Penetrometer"
'
    If H2OPERM = True Then
        If g_bBalanceNotPenet = True Then
            Penetro.Caption = ts$(8) + ": "         '"Penetrometer"
        Else
            Penetro.Caption = ts$(33) + ": "         '"Mettler Balance"
        End If
    End If
'
' END code inserted by Tim Richards 6/4/04
' **********

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
        If bubbler_enable = True Then
            Move_Valve 24, "C" ' close bubbler diverter
        End If
        If slurry_tube_exists Then '6.71.123.19
            Move_Valve Slurry_tube_vent_valve, "C" 'slurry tube vent valve V29
        End If
        If H2OPERM Then
            Move_Valve 8, "C"
            If Drain12 Then
                Move_Valve 11, "C"
            End If
        End If
        If multiChamberSystem Then
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
            move_piston "C"
        End If
        If Num_Microflow_Volumes <= 1 Then
            'volumeSelectHead.Visible = False
        Else
            If Num_Microflow_Volumes < 3 Then
                selvolume(3).Visible = False
            End If
            If microFlowUseAllVolumes Then
                update_microflow_volume_selection_menu (4)
            Else
                update_microflow_volume_selection_menu (Current_Microflow_Volume_Index)
            End If
        End If
    Else
        ' can't change microflow volumes in the middle of a test
        'volumeSelectHead.Visible = False
        ' we are running a test
        If multiChamberSystem Then
            ' set the chambers check box to reflect
            ' which unit is current
            For i% = 1 To chambers
                If Vpos(1 - i%) = 1 Then
                    'chamberiso(i% - 1).value = 1
                    ' when we set the value, it will disable
                    ' the button, so we need to re-enable it
                    'chamberiso(i% - 1).Enabled = True
                End If
            Next i%
            ' if we set any value, it would add a command
            ' to the pending string, so clear it just
            ' in case
            pending$ = ""
        End If
    End If
    If Vpos(15) = 1 Then piston(1).Visible = True
    If dual_stage_compression And Vpos(15) = 0 And Vpos(24) = 1 Then piston(1).Height = 255
    RUNNING = False

    'Form_Paint
    If dry_chambers > 1 Then
        make_valve_visible 19
        If Vpos(5) <> 0 Then
            show_valve_open 19
        End If
    End If
    If DiffPG Then
        'upgclik.Visible = True
        'uflowtubing(0).Visible = True
        'uflowtubing(1).Visible = True
        'upglabel.Visible = True
        'upgshape(0).Visible = True
        'upgshape(1).Visible = True
        ' deal with frazierpressuregauge
        If FrazierPressureGauge Then
            'upglabel.Caption = "Frazier Pressure"
            If FrazierChamberValve > 0 Then
                ValveLabel(14).Caption = "VF"
                make_valve_visible 14
                If Vpos(FrazierChamberValve) <> 0 Then
                    show_valve_open 14
                End If
            End If
        Else
            make_valve_visible 14
            If Vpos(14) <> 0 Then
                show_valve_open 14
            End If
        End If
    End If
    If bubbler_enable Then
        make_valve_visible 18
        If Vpos(25) <> 0 Then
            show_valve_open 18
        End If
        'If BubblerLevelChannel >= 0 Then BubblerLabel.Visible = True
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
        'liqpermshape.Visible = True
        'liquidlabel.Visible = True
        'liquidtubing(0).Visible = True
        'liquidtubing(1).Visible = True
        'liquidtubing(2).Visible = True
        'liquidtubing(3).Visible = True
        'liquidtubing(4).Visible = True
        If Drain12 Then
            'drain12tubing.Visible = True
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
    
'edited 12/10/07 --Denis
    'standards that show all but the auto fill controls
    'ManualControl1.Width = 10000
    'ManualControl1.Height = 8250
    
    If auto_soak_enable And Not hasMultipleMVs Then
'        AutoWet.Visible = True      'if the hardware feature is enabled then show controls
        ' check boxes on screen default to value 0 at design time
        ' no need to initialize them in the form load
'        Move_Valve Fill_ValveA, "C"         'Make sure the valve is closed to match the screen
'        Move_Valve Drain_ValveA, "C"        'Make sure the valve is closed to match the screen
        ' also for chamber 2
'        Move_Valve Fill_ValveB, "C"         'Make sure the valve is closed to match the screen
'        Move_Valve Drain_ValveB, "C"        'Make sure the valve is closed to match the screen
'        Command6.Visible = True
'        Command7.Visible = True
        
    Else
        'AutoWet.Visible = False
    End If
'/////////////////////////////////////////////////////


    ' stuff for valved regulator
    If (Not newreg) And (Not ip_reg_enable) Then
        For i% = 5 To 8
            make_valve_visible i%
        Next i%
        For i% = 0 To 7
            'reg53tubing(i%).Visible = True
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
        'lpgclik.Visible = True
        'lpglabel.Visible = True
        'lpgshape(0).Visible = True
        'lpgshape(1).Visible = True
        'lpgtubing(0).Visible = True
        'lpgtubing(1).Visible = True
    End If
    If xhflow Then
        mflowclik.Visible = True
        mflowlabel.Visible = True
        mflowshape(0).Visible = True
        mflowshape(1).Visible = True
        mftubing(0).Visible = True
        mftubing(1).Visible = True
        newmflowclik.Visible = True
        newmflowlabel.Visible = True
        mflowshape(2).Visible = True
        mflowshape(3).Visible = True
        newmftubing(0).Visible = True
        newmftubing(1).Visible = True
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
        mv1click(0).Visible = False
        mv1click(1).Visible = False
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
    If Auto_fill And version < 7 Then
        make_valve_visible 0
        'botfilltubing.Visible = True
    End If
    If Auto_fill And version >= 7 Then
        ' version 7 autofill is a top fill valve
        make_valve_visible 13
        'topfilltubing(0).Visible = True
        'topfilltubing(1).Visible = True
    End If
    If integrity Then
        Rem integrity flow meter added
        'iflowclik.Visible = True
        'iflowlabel.Visible = True
        'iflowshape(0).Visible = True
        'iflowshape(1).Visible = True
        'iflowtubing(0).Visible = True
        'iflowtubing(1).Visible = True
    End If
    If integrity Or DiffPG Then
        If AirTop Then
            'auxbottubing(0).Visible = True
            'auxbottubing(1).Visible = True
            'auxbottubing(2).Visible = True
        Else
            'auxtoptubing.Visible = True
        End If
    End If
    If GasPerm Then ' gasperm is true if there is no low flow meter
        ' hide low flow meter, v1 tubing, don't show valve 1
        'lflowclik.Visible = False
        'lflowlabel.Visible = False
        'lflowshape(0).Visible = False
        'lflowshape(1).Visible = False
        'fixedneedleshape(0).Visible = False
        'fixedneedleshape(1).Visible = False
        
        'v1tubing(0).Visible = False
        'v1tubing(1).Visible = False
        'v1tubing(2).Visible = False
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
        'airtoptubing(2).Visible = True
        If v20_exists Then
            make_valve_visible 2
            If Vpos(20) <> 0 Then
                show_valve_open 2
            End If
        End If
    ElseIf liqpermonly Then
        ' valve picture 2 is used for valve 12 on liqpermonly
'        make_valve_visible 2
'        ValveLabel(2).Caption = ts$(14)     ' "Isolation"
'        If Vpos(12) <> 0 Then
'            show_valve_open 2
'        End If
        
        ' show manual drain valve
        'mandrainlabel.Visible = True
        'mandrainshape.Visible = True
    Else
        'airbottubing(0).Visible = True
        'airbottubing(1).Visible = True
        'airbottubing(2).Visible = True
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
    If dryChamberTemperature <> 0 Or wetChamberTemperature <> 0 Or reservoirTemperature <> 0 _
       Or airTemperature <> 0 Or bubblerTemperature <> 0 Or cabinetTemperature <> 0 _
       Or hydroHeadTemperature <> 0 Or mullenTemperature <> 0 Then
        'templabel.Visible = True
        'tempset.Visible = True
        'tempval.Visible = True
        'Combo1.Visible = True
        If reservoirTemperature <> 0 Then
            'Combo1.AddItem reservoirLabel
        End If
        If dryChamberTemperature <> 0 Then
            'Combo1.AddItem dryprobeLabel
        End If
        If wetChamberTemperature <> 0 Then
            'Combo1.AddItem wetprobeLabel
        End If
        If airTemperature <> 0 Then
            'Combo1.AddItem airprobeLabel
        End If
        If bubblerTemperature <> 0 Then
            'Combo1.AddItem bubblerprobeLabel
        End If
        If cabinetTemperature <> 0 Then
            'Combo1.AddItem cabinetprobeLabel
        End If
        If hydroHeadTemperature <> 0 Then
            'Combo1.AddItem hydroheadprobeLabel
        End If
        If mullenTemperature <> 0 Then
            'Combo1.AddItem mullenprobeLabel
        End If
        'Combo1.ListIndex = 0
    End If
    TitleScrn.MousePointer = 0
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(15) + str$(unitnumber)    ' "Unit"
    End If
    If recirculation Then
        make_valve_visible 15
        If Vpos(21) <> 0 Then
            show_valve_open 15
        End If
        If v22_exists Then
            'Me.recircline1.Visible = True
            make_valve_visible 16
            If Vpos(22) <> 0 Then
                show_valve_open 16
            End If
        End If
    End If
    If ReserveTankLevelChannel >= 0 Then
        lblReserveTankLevel.Visible = True
        lblReserveTankLevelValue.Visible = True
    Else
        lblReserveTankLevel.Visible = False
        lblReserveTankLevelValue.Visible = False
    End If
    If valve_23_exists Then
        make_valve_visible 17
        If Vpos(23) <> 0 Then
            show_valve_open 17
        End If
    End If
    
    If sampleChamberDiverterValve >= 0 Then
        ValveLabel(27).Visible = True
        ValveClik(27).Visible = True
        vopenline(27).Visible = True
        vcloseline(27).Visible = True
        ValveFill(27).Visible = True
        ValveLabel(27).Visible = True
        If Vpos(sampleChamberDiverterValve + 1) = 0 Then
            show_valve_open 27
        Else
            show_valve_closed 27
        End If
    Else
        ValveLabel(27).Visible = False
        ValveClik(27).Visible = False
        vopenline(27).Visible = False
        vcloseline(27).Visible = False
        ValveFill(27).Visible = False
        ValveLabel(27).Visible = False
    End If
    
    If number_of_wetting_valves > 0 Then
        WettingFrame.Visible = True
        chkWettingValve1.ToolTipText = "Valve " & str$(wetting_valve(1))
    Else
        WettingFrame.Visible = False
    End If
    
    'If doorlock Then Command5.Visible = True
    
    ' Visibility of autocal button
    Ret$ = String$(255, 0)
    GPPS "Capstuff", "CalibrateWindow", "", Ret$, 255, CSFile$
    a$ = nulltrim(Ret$)
    'Command4.Visible = (a$ = "Y")
    
    ' initialize future variables.  -1 means no current request
    future_pres% = -1
    future_lflow% = -1
    future_hflow% = -1
    future_DPress% = -1
    future_CPress% = -1

' **********
' BEGIN code inserted by search for Tim Richards Friday 6/4/04
' change the form if using a balance
'
    Dim iVFP_offset_X As Integer
    Dim iVFP_offset_Y As Integer
    Dim i_X1, i_Y1, i_X2, i_Y2 As Integer

    'ValveDrain13_PlaceHolder.Visible = False

    If g_bBalanceNotPenet = True Then

        ' Hide the penetrometer
        'liqpermshape.Visible = False
        'topfilltubing(0).Visible = False
        'topfilltubing(1).Visible = False
        'liquidlabel.Visible = False

        ' Score the difference between the old and new
        'iVFP_offset_X = ValveDrain13_PlaceHolder.Left - ValveFill(13).Left
        'iVFP_offset_Y = ValveDrain13_PlaceHolder.top - ValveFill(13).top

        ' Move the Penetrometer Fill valve & rename it
        'ValveFill(13).Left = ValveDrain13_PlaceHolder.Left
        'ValveFill(13).top = ValveDrain13_PlaceHolder.top
        vopenline(13).x1 = vopenline(13).x1 + iVFP_offset_X
        vopenline(13).x2 = vopenline(13).x2 + iVFP_offset_X
        vopenline(13).y1 = vopenline(13).y1 + iVFP_offset_Y
        vopenline(13).Y2 = vopenline(13).Y2 + iVFP_offset_Y
        vcloseline(13).x1 = vcloseline(13).x1 + iVFP_offset_X
        vcloseline(13).x2 = vcloseline(13).x2 + iVFP_offset_X
        vcloseline(13).y1 = vcloseline(13).y1 + iVFP_offset_Y
        vcloseline(13).Y2 = vcloseline(13).Y2 + iVFP_offset_Y
        ValveLabel(13).Left = ValveLabel(13).Left + iVFP_offset_X
        ValveLabel(13).top = ValveLabel(13).top + iVFP_offset_Y
        'ValveClik(13).Left = ValveDrain13_PlaceHolder.Left
        'ValveClik(13).top = ValveDrain13_PlaceHolder.top

        ' Score the difference between the old and new
        'iVFP_offset_X = ValveIso12_PlaceHolder.Left - ValveFill(12).Left
        'iVFP_offset_Y = ValveIso12_PlaceHolder.top - ValveFill(12).top

        ' Move the Drain valve and turn it into an isolation valve.
        ' Use temp to swap the lines
        'ValveFill(12).Left = ValveIso12_PlaceHolder.Left
        'ValveFill(12).top = ValveIso12_PlaceHolder.top
        vopenline(12).x1 = vopenline(12).x1 + iVFP_offset_X
        vopenline(12).x2 = vopenline(12).x2 + iVFP_offset_X
        vopenline(12).y1 = vopenline(12).y1 + iVFP_offset_Y
        vopenline(12).Y2 = vopenline(12).Y2 + iVFP_offset_Y
        vcloseline(12).x1 = vcloseline(12).x1 + iVFP_offset_X
        vcloseline(12).x2 = vcloseline(12).x2 + iVFP_offset_X
        vcloseline(12).y1 = vcloseline(12).y1 + iVFP_offset_Y
        vcloseline(12).Y2 = vcloseline(12).Y2 + iVFP_offset_Y
        ValveLabel(12).Left = ValveLabel(12).Left + iVFP_offset_X
        ValveLabel(12).top = ValveLabel(12).top + iVFP_offset_Y
        'ValveClik(12).Left = ValveIso12_PlaceHolder.Left
        'ValveClik(12).top = ValveIso12_PlaceHolder.top
        'drain12tubing.Visible = False

        'use templine to swap the drain and fill lines
        i_X1 = vopenline(12).x1
        i_Y1 = vopenline(12).y1
        i_X2 = vopenline(12).x2
        i_Y2 = vopenline(12).Y2
        vopenline(12).x1 = vcloseline(12).x1
        vopenline(12).y1 = vcloseline(12).y1
        vopenline(12).x2 = vcloseline(12).x2
        vopenline(12).Y2 = vcloseline(12).Y2
        vcloseline(12).x1 = i_X1
        vcloseline(12).y1 = i_Y1
        vcloseline(12).x2 = i_X2
        vcloseline(12).Y2 = i_Y2
        'label_Iso.Visible = True
        
        'globally localize the labellies
        'label_MettlerBalance.Caption = ts$(34)
        'label_Iso.Caption = ts$(35)
        ValveLabel(13).Caption = ts$(36)

    Else

        ' Hide the balance
        'shape_MettlerBalance.Visible = False
        'line_MettlerBalance.Visible = False
        'line_MettlerDrain.Visible = False
        'label_MettlerBalance.Visible = False

    End If
'
' END code inserted by Tim Richards 6/4/04
' **********
    piston_position_transducer.Visible = piston_position_transducer_exists '6.71.123.08
    'slurry_tube_level.Visible = slurry_tube_exists '6.71.123.08
    'slurry_tube_frame.Visible = slurry_tube_exists '6.71.123.08
    'slurry_tube_pressure.Visible = slurry_tube_exists '6.71.123.08
    'slurry_wash_pump_flow.Visible = slurry_tube_exists '6.71.123.10
    If slurry_tube_exists Then                         '6.71.123.10
        update_slurry_wash_pump_flow_display
    End If

    If num_sample_pressure_gauges > 0 Then
        piston_position_transducer.Visible = False
        Auxreading.Visible = False
        Penetro.top = 1610
        
        make_valve_visible 5
        If Vpos(5) <> 0 Then
            show_valve_open 5
        End If
        
        make_valve_visible 6
        
        If Vpos(6) <> 0 Then
            show_valve_open 6
        End If
        
        If num_sample_pressure_gauges = 1 Then
            'SamplePG1Label.Visible = True
            'SamplePG1Label.top = 1900
            'SamplePG2Label.Visible = False
            
        ElseIf num_sample_pressure_gauges = 2 Then
            'SamplePG1Label.Visible = True
            'SamplePG1Label.top = 1900
            'SamplePG2Label.Visible = True
            'SamplePG2Label.top = 2200
        Else
            'SamplePG1Label.Visible = False
            'SamplePG2Label.Visible = False
        End If
    Else
        'SamplePG1Label.Visible = False
        'SamplePG2Label.Visible = False
    End If
    
    If tank_level_exists Then
        'tankLabel.Visible = True
        'tankLabel.BackStyle = 1
        'tankLabel.ForeColor = &H80000012
        'tankLevelLabel.Visible = True
        'tankLevelLabel.BackStyle = 1
        'tankLevelLabel.ForeColor = &H80000012
    End If
    
    'AJB 11-06-09
    mainmenu(7).Visible = test_piston

    'AJB 12-14-09
    If hasMultipleMVs Then
        mainmenu(8).Visible = True
    Else
        motorValveIndex = -1
        mainmenu(8).Visible = False
    End If
        
    'AJB 12-23-09 new pump / wettign valve controls
    If number_of_wetting_valves > 0 Then
        cmdShowWettingControls.Visible = True
    End If
    
    If hasHumidityControls Then
        make_valve_visible (18)
        Move_Valve 3, "O"
    End If
    
    'rvw 4-23-10 Resin Intrusion controls
    If Resin_Diverter_Valve = 0 Then
        'ResinFrame.Visible = False
    End If
    
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
End Sub

'//////////////////////////////////
'edited 10/11/07 --Denis
Private Sub Fill_Click()
    'Fill.Enabled = False
    do_command 70
End Sub

Private Sub Drain_Click()
    'Drain.Enabled = False
    do_command 71
End Sub

Private Sub Fill2_Click()
    'Fill2.Enabled = False
    do_command 72
End Sub

Private Sub Drain2_Click()
    'Drain2.Enabled = False
    do_command 73
End Sub
'//////////////////////////////////////


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
        If version = 6 Then
            msg$ = msg$ + vbCrLf + ts$(23)      ' "Shift key multiplies command by 10"
        End If
        MsgBox msg$, 64, ts$(6) + " " + ts$(7)  ' "Pressure Regulator"
    Case 4:
        msg$ = ts$(24) + vbCrLf + ts$(25) + vbCrLf + ts$(26)       ' "Click on left half to pulse open."/"Click on right half to pulse closed."/"Press S, O, or C to STOP, OPEN, or CLOSE fully."
        MsgBox msg$, 64, ts$(30)                ' "Needle Valve 2 ONLY"
    End Select
Exit Sub

End Sub

Private Sub mainmenu_Click(Index As Integer)

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
        Case 7
            do_command 84
        Case 8
            do_command 89
    End Select
    
End Sub

Private Sub penselect_Click(Index As Integer)

    If penetrometer_select <> Index + 1 Then
        Rem change penetrometers
        do_command 59
    End If
    
End Sub

Private Sub mv4click_Click(Index As Integer)
    If motorValveIndex <> 2 Then
        motorValveIndex = 2
        do_command 55
    End If
    do_command 10 + Index
End Sub

Private Sub mv2click_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        motorValveIndex = 1
        do_command 55
    ElseIf Button = 2 Then
        If motorValveIndex <> 1 Then
            motorValveIndex = 1
            do_command 55
        End If
        do_command 10 + Index
    End If

End Sub

Private Sub mv3click_Click(Index As Integer)
    If motorValveIndex <> 2 Then
        motorValveIndex = 2
        do_command 55
    End If
    do_command 10 + Index
End Sub

Private Sub mv3click_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        motorValveIndex = 2
        do_command 55
    ElseIf Button = 2 Then
        If motorValveIndex <> 2 Then
            motorValveIndex = 2
            do_command 55
        End If
        do_command 10 + Index
    End If
End Sub

Private Sub newhflowclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    motorValveIndex = 1
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
                ' show valve 10' open
                ' only do this for valves that are actually visible
                If newValveFill.Visible = False Then Exit Sub
                newvcloseline.Visible = False
                newvopenline.Visible = True
                ' now for color
                If newValveFill.FillColor = vbGreen Or newValveFill.FillColor = vbRed Then
                    newValveFill.FillColor = vbGreen
                Else
                    newValveFill.FillColor = vbYellow
                End If
            End If
            ' 6.71.20 end
        End If
        auto_index% = 2
End Sub

Private Sub newmflowclik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    motorValveIndex = 1
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
        
        If newValveFill.Visible = False Then Exit Sub
        newvcloseline.Visible = True
        newvopenline.Visible = False
        ' now for color
        If newValveFill.FillColor = vbGreen Or newValveFill.FillColor = vbRed Then
            newValveFill.FillColor = vbRed
        Else
            newValveFill.FillColor = vbCyan
        End If
        
    End If
    ' 6.71.20 end
    auto_index% = 2
    
End Sub

Private Sub newv2clik_Click(Index As Integer)
    motorValveIndex = 1
    do_command 55
    do_command 10 + Index
End Sub

Private Sub pistonclik_Click(Index As Integer)
    do_command 53 ' piston toggle
End Sub

Private Sub pistonclik2_Click(Index As Integer)
    do_command 74 ' 2nd piston toggle
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
    'regrew.Enabled = False
    'regff.Enabled = False
    do_command 44
End Sub

Private Sub regrew_Click()
    'regrew.Enabled = False
    'regff.Enabled = False
    do_command 42
End Sub

Private Sub RegScroll_Change()
    reglabel.Caption = ts$(12) + ":" + str$(RegScroll.value)    ' "Reg Jump"
End Sub

Private Sub regsel_Click(Index As Integer)
    do_command 54 + Index
End Sub

Private Sub regstop_Click()
    'regstop.Enabled = False
    do_command 43
End Sub

Private Sub ResinVacuumCheckBox_Click()
    'ResinVacuumCheckBox.Enabled = False
    pending$ = pending$ + Chr$(96)
End Sub

Private Sub ResinValveCheckBox_Click()
    'ResinValveCheckBox.Enabled = False
    pending$ = pending$ + Chr$(95)
End Sub

Private Sub selvolume_Click(Index As Integer)
    update_microflow_volume_selection_menu (Index)
End Sub

Private Sub SlurryPumpSpeed_Click(Index As Integer) '6.71.123.10
    do_command 81 + Index
End Sub

Private Sub SlurryPumpSpeedScroll_Change() '6.71.123.10
    'SlurryPumpSpeedLabel.Caption = "Count Jump:" + str$(SlurryPumpSpeedScroll.value)
End Sub

Private Sub statusmenu_Click(Index As Integer)
' 0=off, 1=red, 2=yellow
status_lights_value = Index
do_command 69
End Sub

Private Sub templabel_Click()
' this label acts as the temperature selector if there is only one probe
' it acts this way all the time for now
'If temperature% = 1 Then
    manual_aux_click = 1 ' signify temperature was last clicked on
'End If
End Sub

'Private Sub tempprobe1_Click()
'manual_aux_click = 1 ' signify temperature was last clicked on
'End Sub

'Private Sub tempprobe2_Click()
'manual_aux_click = 1 ' signify temperature was last clicked on
'End Sub

Private Sub tempset_Click()

    Dim temptarget%

    'temptarget% = myVal(tempval.Text) * 10
    'If temptarget% < 0 Then temptarget% = 0
    'If temptarget% > 9999 Then temptarget% = 9999
    'tempval.Text = str$(temptarget% / 10#)
    'pending$ = pending$ + Chr$(45) + Chr$(temptarget% And 127) + Chr$(Int(temptarget% / 128))

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
        Me.Caption = Me.Caption + " - " + ts$(15) + str$(unitnumber)    ' "Unit"
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
    If pulseV1Enable Then
        do_command 1
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
        Case 4
            pulseV1Enable = Not pulseV1Enable
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
    If hasHumidityControls And Index = 18 Then
        do_command 4
    Else
        ' 10 and 11 are not clickable, others map to same commmand
        ' 15 and 16 and 17 and 18 map to different commands
        ' 19 maps to new second dry chamber valve 5
        If Index = 10 Or Index = 11 Then
            Exit Sub
        ElseIf Index = 15 Or Index = 16 Or Index = 17 Then
            do_command Index + 2 '(15=17, 16=18, 17=19)
        ElseIf Index = 18 Then
            do_command 65 ' 18 is valve 25 which is command 65 (confused yet?) - new bubbler valve
        ElseIf Index = 19 Then
            do_command 66 ' 19 is new valve 5 which is command 66
        ElseIf Index >= 20 And Index <= 25 Then   '6.71.123.03
            do_command Index + 55 '(20=75, 21=76, 22=77, 23=78, 24=79, 25=80)
        ElseIf Index = 27 Then
            do_command 101
        Else
            do_command Index
        End If
    End If

End Sub

Private Sub xhhvalve_Click(Index As Integer)
    'xhhvalve(Index).Enabled = False
    'pending$ = pending$ + Chr$(56 + Index)
End Sub
Private Sub make_valve_visible(i As Integer)
    If i = 1 Or i = 3 Or i = 10 Or i = 27 Then
    
        ValveClik(i).Visible = True
        ValveFill(i).Visible = True
        ValveLabel(i).Visible = True
        vcloseline(i).Visible = True
        ' leave the open line hidden
        If i = 10 Then
            newValveClik.Visible = True
            newValveFill.Visible = True
            newValveLabel.Visible = True
            newvcloseline.Visible = True
        End If
    End If
End Sub
Sub show_valve_open(i As Integer)
    If i = 1 Or i = 3 Or i = 10 Or i = 27 Then
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
    End If
End Sub
Sub show_valve_closed(i As Integer)
    If i = 1 Or i = 3 Or i = 10 Or i = 27 Then
        If ValveFill(i).Visible = False Then Exit Sub
        vcloseline(i).Visible = True
        vopenline(i).Visible = False
        ' now for color
        If ValveFill(i).FillColor = vbGreen Or ValveFill(i).FillColor = vbRed Then
            ValveFill(i).FillColor = vbRed
        Else
            ValveFill(i).FillColor = vbCyan
        End If
    End If
End Sub
Sub set_liqtubing(b As Boolean)

    ' true = liqtop
    ' false=liqbot
    'liqbottubing(0).Visible = Not b
    'liqbottubing(1).Visible = Not b
    'liqtoptubing(0).Visible = b
    'liqtoptubing(1).Visible = b
    'liqtoptubing(2).Visible = b

End Sub

Public Sub update_microflow_volume_selection_menu(i As Integer)
    selvolume(1).Checked = False
    selvolume(2).Checked = False
    selvolume(3).Checked = False
    selvolume(4).Checked = False
    selvolume(i).Checked = True
    ' calling the titlescreen version of this will set the global variable
    ' and write the results to the ini file
   ' TitleScrn.update_microflow_volume_selection_menu (i)
    ' this version also needs to update the valves since we are actually in control of the system at this time
    update_microflow_volume_valves
End Sub

Public Sub updateBubblerMVPicture()
    Dim temp_l_value%
    temp_l_value% = bubblerMV_plunger_left% - Int(8 * x5 + 0.5) * 15
    'If ManualControl1.BubblerMVShape(0).Left <> temp_l_value% Then
        'ManualControl1.BubblerMVShape(0).Left = temp_l_value%
    'End If
End Sub

Public Sub LoadTextStrings()
' Load text elements for this form from external .ini file
    
    Dim i As Integer
    ' Manual control screen uses extra fonts for some of its labels, represented here:
    Dim mclabel_font As font_info, mc2_font As font_info
    
    ' Load extra font used for labels. Currently no way to set this font
    ' within the program, but that can be dealt with later if necessary.
    mclabel_font.font = gpps2("default", "mclabel_fontname", language$, "Arial")
    mclabel_font.fontsize = val(gpps2("default", "mclabel_fontsize", language$, "8"))
    mclabel_font.fontbold = (gpps2("default", "mclabel_fontbold", language$, "N") = "Y")
    mc2_font.font = gpps2("default", "mc2_fontname", language$, "MS Serif")
    mc2_font.fontsize = val(gpps2("default", "mc2_fontsize", language$, "7"))
    mc2_font.fontbold = (gpps2("default", "mc2_fontbold", language$, "N") = "Y")
    
    ' Form elements
    ' Window title is set at the end
    For i = 1 To 5
        mainmenu(i).Caption = gpps2("manual", "mainmenu" + str$(i), language$, mainmenu(i).Caption)
    Next i
    For i = 1 To 3
        TMenu(i).Caption = gpps2("manual", "tmenu" + str$(i), language$, TMenu(i).Caption)
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
    'highflow.Caption = get_thing("manual", "highflow", language$, highflow.Caption, highflow, default_font)
    'iflowlabel.Caption = get_thing("manual", "iflow", language$, iflowlabel.Caption, iflowlabel, mclabel_font)
    lfclabel(0).Caption = get_thing("manual", "lfclabel(0)", language$, lfclabel(0).Caption, lfclabel(0), mclabel_font)
    For i = 0 To 2
        lfctrl(i).Caption = gpps2("manual", "lfctrl" + str$(i), language$, lfctrl(i).Caption)
    Next i
    'lflowlabel.Caption = get_thing("manual", "lflowlabel", language$, lflowlabel.Caption, lflowlabel, mclabel_font)
    'liquidlabel.Caption = get_thing("manual", "liquidlabel", language$, liquidlabel.Caption, liquidlabel, mclabel_font)
    'lowflow.Caption = get_thing("manual", "lowflow", language$, lowflow.Caption, lowflow, default_font)
    'mandrainlabel.Caption = get_thing("manual", "mandrain", language$, mandrainlabel.Caption, mandrainlabel, mclabel_font)
    mflowlabel.Caption = get_thing("manual", "mflowlabel", language$, mflowlabel.Caption, mflowlabel, mclabel_font)
    'penselect(0).Caption = get_thing("manual", "penselect0", language$, penselect(0).Caption, penselect(0), mc2_font)
    'penselect(1).Caption = get_thing("manual", "penselect1", language$, penselect(1).Caption, penselect(1), mc2_font)
    pistonlabel.Caption = get_thing("manual", "pistonlabel", language$, pistonlabel.Caption, pistonlabel, mclabel_font)
    Press_Read.Caption = get_thing("manual", "press_read", language$, Press_Read.Caption, Press_Read, default_font)
    reglabel.Caption = get_thing("manual", "reglabel", language$, reglabel.Caption, reglabel, mclabel_font)
    Regulator.Caption = get_thing("manual", "regulator", language$, Regulator.Caption, Regulator, default_font)
    'templabel.Caption = get_thing("manual", "templabel", language$, templabel.Caption, templabel, default_font)
    'tempprobe1.Caption = get_thing("manual", "tempprobe1", language$, tempprobe1.Caption, tempprobe1, default_font)
    'tempprobe2.Caption = get_thing("manual", "tempprobe2", language$, tempprobe2.Caption, tempprobe2, default_font)
    'tempset.Caption = gpps2("manual", "tempset", language$, tempset.Caption)
    'set_fontstuff tempval, default_font
    set_fontstuff TLabel, default_font
    'upglabel.Caption = get_thing("manual", "upglabel", language$, upglabel.Caption, upglabel, mclabel_font)
    Valve_Pos(2).Caption = get_thing("manual", "valvepos2", language$, Valve_Pos(2).Caption, Valve_Pos(2), default_font)
    'ValveLabel(0).Caption = get_thing("manual", "valvelabel0", language$, ValveLabel(0).Caption, ValveLabel(0), mc2_font)
    'ValveLabel(2).Caption = get_thing("manual", "valvelabel2", language$, ValveLabel(2).Caption, ValveLabel(2), mc2_font)
    'ValveLabel(3).Caption = get_thing("manual", "valvelabel3", language$, ValveLabel(3).Caption, ValveLabel(3), mc2_font)
    'ValveLabel(12).Caption = get_thing("manual", "valvelabel12", language$, ValveLabel(12).Caption, ValveLabel(12), mc2_font)
    'ValveLabel(13).Caption = get_thing("manual", "valvelabel13", language$, ValveLabel(13).Caption, ValveLabel(13), mc2_font)
    'xhhtitle.Caption = get_thing("manual", "xhhtitle", language$, xhhtitle.Caption, xhhtitle, mc2_font)
    For i = 0 To 2
        'xhhvalve(i).Caption = get_thing("manual", "xhhvalve" + str$(i), language$, xhhvalve(i).Caption, xhhvalve(i), mc2_font)
    Next i
    Penetro.Caption = get_thing("manual", "penetro", language$, Penetro.Caption, Penetro, default_font)
    Auxreading.Caption = get_thing("manual", "auxreading", language$, Auxreading.Caption, Auxreading, default_font)
    set_fontname Command1, default_font
    'set_fontname Command2, default_font
    'set_fontname Command3, default_font
    'set_fontname Command4, default_font
    'set_fontname Command5, default_font
    'Command4.Caption = gpps2("manual", "autocal", language$, Command4.Caption)
    Command1.Caption = gpps2("manual", "command1", language$, Command1.Caption)
    'Command2.Caption = gpps2("manual", "command2", language$, Command2.Caption)
    'Command3.Caption = gpps2("manual", "command3", language$, Command3.Caption)
    'Command5.Caption = gpps2("manual", "command5", language$, Command5.Caption)
    
    'reservoirLabel = get_thing("manual", "reservoirprobe", language$, "reservoir", Combo1, default_font)
    wetprobeLabel = gpps2("manual", "wetprobe", language$, "wet")
    dryprobeLabel = gpps2("manual", "dryprobe", language$, "dry")
    airprobeLabel = gpps2("manual", "airprobe", language$, "air")
    bubblerprobeLabel = gpps2("manual", "bubblerprobe", language$, "bubbler")
    cabinetprobeLabel = gpps2("manual", "cabinetprobe", language$, "cabinet")
    If burst Then
        hydroheadprobeLabel = "burst"
    Else
        hydroheadprobeLabel = gpps2("manual", "hydroheadprobe", language$, "hydrohead")
    End If
    mullenprobeLabel = gpps2("manual", "mullenprobe", language$, "mullen")
    
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
    
' **********
' Begin localization code entered by search for Tim Richards on Tuesday 6/8/04
'
    ts$(33) = gpps2("BalanceNotPenet", "ts33", language$, "Mettler Balance")
    ts$(34) = gpps2("BalanceNotPenet", "ts34", language$, "Mettler Balance")
    ts$(35) = gpps2("BalanceNotPenet", "ts35", language$, "Iso V12")
    ts$(36) = gpps2("BalanceNotPenet", "ts36", language$, "Drain V13")
'
' End code entered by Tim Richards 6/8/04
' **********
    
    ManualControl1.Caption = get_thing("manual", "window title", language$, ts$(31), ManualControl1, default_font)

End Sub



