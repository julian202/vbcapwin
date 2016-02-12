VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form TitleScrn 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automated Capillary Flow Porometer"
   ClientHeight    =   8400
   ClientLeft      =   390
   ClientTop       =   795
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "CAPMAIN.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8400
   ScaleWidth      =   8415
   Begin VB.Frame Frame1 
      BackColor       =   &H0000B3FE&
      BorderStyle     =   0  'None
      Height          =   8145
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8175
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   6240
         Picture         =   "CAPMAIN.frx":030A
         ScaleHeight     =   600
         ScaleWidth      =   1095
         TabIndex        =   24
         Top             =   960
         Width           =   1095
      End
      Begin VB.Frame GasFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gas"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   3120
         TabIndex        =   13
         Top             =   6000
         Width           =   3135
         Begin VB.Label GasLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Label4"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame PressUnitFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pressure Units"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   3120
         TabIndex        =   9
         Top             =   4200
         Width           =   3135
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Label6"
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   2925
         End
      End
      Begin VB.Frame lohmTableFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lohm Table"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   3120
         TabIndex        =   12
         Top             =   5400
         Width           =   3135
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Label1"
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   2895
         End
      End
      Begin VB.Frame GroupNmFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Group Name"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   3120
         TabIndex        =   8
         Top             =   4800
         Width           =   3135
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "SID: "
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   2925
         End
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   7935
         Left            =   120
         Picture         =   "CAPMAIN.frx":25AC
         ScaleHeight     =   7875
         ScaleWidth      =   7875
         TabIndex        =   17
         Top             =   120
         Width           =   7935
         Begin VB.CommandButton Preferences 
            BackColor       =   &H000CCAFA&
            Caption         =   "Preferences"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6480
            Width           =   1575
         End
         Begin VB.CommandButton parametersCMD 
            BackColor       =   &H000CCAFA&
            Caption         =   "Auto Parameters"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   5880
            Width           =   1575
         End
         Begin VB.CommandButton executeReportCMD 
            BackColor       =   &H000CCAFA&
            Caption         =   "Open Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   5280
            Width           =   1575
         End
         Begin VB.CommandButton manualTestCMD 
            BackColor       =   &H000CCAFA&
            Caption         =   "Manual Control"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   4680
            Width           =   1575
         End
         Begin VB.CommandButton autoTestCMD 
            BackColor       =   &H000CCAFA&
            Caption         =   "Auto Test"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   4080
            Width           =   1575
         End
         Begin VB.Label experimental2Label 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "EXPERIMENTAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   23
            Top             =   720
            Width           =   2655
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   480
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BackColor       =   &H00D0D0D0&
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   600
         Picture         =   "CAPMAIN.frx":22FAD
         ScaleHeight     =   690
         ScaleWidth      =   7020
         TabIndex        =   2
         Top             =   360
         Width           =   7020
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Call CalcMFP"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7320
         Top             =   4920
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Call CalcGP"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1950
         Left            =   1920
         Picture         =   "CAPMAIN.frx":27AAA
         ScaleHeight     =   1950
         ScaleWidth      =   4245
         TabIndex        =   0
         Top             =   1920
         Width           =   4245
      End
      Begin MSCommLib.MSComm AuxComm 
         Left            =   0
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   2
         DTREnable       =   0   'False
         RThreshold      =   1
      End
      Begin MSCommLib.MSComm MainComm 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InputLen        =   1
         ParityReplace   =   32
         RThreshold      =   1
         InputMode       =   1
      End
      Begin VB.Label lblUsername 
         Alignment       =   2  'Center
         BackColor       =   &H00D0D0D0&
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   3960
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label experimentalLabel 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   $"CAPMAIN.frx":2C37B
         Height          =   855
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   7215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "Automated Capillary Flow Porometer System Software for Windows"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   960
         TabIndex        =   15
         Top             =   1080
         Width           =   6525
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00D0D0D0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderWidth     =   5
         Height          =   7545
         Left            =   240
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Group"
      Index           =   0
      Begin VB.Menu filemenu 
         Caption         =   "&Select Group"
         Index           =   1
         Shortcut        =   ^U
      End
      Begin VB.Menu filemenu 
         Caption         =   "&New Group"
         Index           =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu filemenu 
         Caption         =   "&Delete Group"
         Index           =   3
      End
      Begin VB.Menu filemenu 
         Caption         =   "&Edit Group Name"
         Index           =   4
      End
      Begin VB.Menu filemenu 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mode 
         Caption         =   "Mode"
         Begin VB.Menu modemenu 
            Caption         =   "Supervisor"
            Index           =   0
         End
         Begin VB.Menu modemenu 
            Caption         =   "User"
            Index           =   1
         End
         Begin VB.Menu modemenu 
            Caption         =   "QC"
            Index           =   2
         End
         Begin VB.Menu modemenu 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu modemenu 
            Caption         =   "Change Supervisor Password"
            Index           =   4
         End
      End
      Begin VB.Menu d1 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLogoutSepLine 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu exitmenu 
         Caption         =   "E&xit"
         Index           =   0
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Execute"
      Index           =   1
      Begin VB.Menu runmenu 
         Caption         =   "&AutoTest"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu runmenu 
         Caption         =   "&Manual Control"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu runmenu 
         Caption         =   "&CleanOut"
         Index           =   3
      End
      Begin VB.Menu runmenu 
         Caption         =   "&Leak Test"
         Index           =   4
         Shortcut        =   ^L
      End
      Begin VB.Menu runmenu 
         Caption         =   "QC Bubble Point Test"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu runmenu 
         Caption         =   "Cyclic Compression"
         Index           =   6
      End
      Begin VB.Menu runmenu 
         Caption         =   "Vacuum Purge"
         Index           =   7
      End
      Begin VB.Menu runmenu 
         Caption         =   "Temperature Log"
         Index           =   8
      End
      Begin VB.Menu chselect 
         Caption         =   "C&hamber Select"
      End
      Begin VB.Menu volumeSelectHead 
         Caption         =   "Microflow Volume Select"
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
      Begin VB.Menu penselecthead 
         Caption         =   "&Penetrometer Select"
         Begin VB.Menu selpen 
            Caption         =   "Penetrometer 1"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu selpen 
            Caption         =   "Penetrometer 2"
            Index           =   2
         End
      End
      Begin VB.Menu mnuMicrometerSelect 
         Caption         =   "Micrometer Select"
         Begin VB.Menu mnuMicrometer 
            Caption         =   "Micrometer 1"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuMicrometer 
            Caption         =   "Micrometer 2"
            Index           =   1
         End
         Begin VB.Menu mnuMicrometer 
            Caption         =   "Micrometer 3"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Report"
      Index           =   2
      Begin VB.Menu reportmenu 
         Caption         =   "Execute &Report"
         Index           =   1
         Shortcut        =   {F4}
      End
      Begin VB.Menu reportmenu 
         Caption         =   "Alter &Cover"
         Index           =   2
      End
      Begin VB.Menu viewReportMenu 
         Caption         =   "View &Report"
         Index           =   3
      End
      Begin VB.Menu convertData 
         Caption         =   "Convert Data"
         Index           =   4
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Modify"
      Index           =   3
      Begin VB.Menu modifymenu 
         Caption         =   "&Edit Auto Parameters"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu modifymenu 
         Caption         =   "Manually Create &File"
         Index           =   1
      End
      Begin VB.Menu modifymenu 
         Caption         =   "&CurveFit"
         Index           =   2
      End
      Begin VB.Menu modifymenu 
         Caption         =   "&Data Editor"
         Index           =   3
      End
      Begin VB.Menu modifymenu 
         Caption         =   "Data &Averager"
         Index           =   4
      End
      Begin VB.Menu psl 
         Caption         =   "Pressure Step Lists"
         Index           =   0
      End
      Begin VB.Menu d1a 
         Caption         =   "Border Color"
         Index           =   0
      End
      Begin VB.Menu prefs 
         Caption         =   "Preferences"
         Index           =   0
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Calibrate"
      Index           =   5
      Begin VB.Menu calibmenu 
         Caption         =   "&Set Valve Limits"
         Index           =   1
      End
      Begin VB.Menu calibmenu 
         Caption         =   "&Adjust Flow Parameters"
         Index           =   2
      End
      Begin VB.Menu calibmenu 
         Caption         =   "&Calculate Lohm and Max Air Flow"
         Index           =   3
      End
      Begin VB.Menu calibmenu 
         Caption         =   "&Max-Liquid Flow"
         Index           =   4
      End
      Begin VB.Menu calibmenu 
         Caption         =   "Micro&Flow Volume"
         Index           =   5
      End
      Begin VB.Menu calibmenu 
         Caption         =   "Calibrate &Regulator"
         Index           =   6
      End
      Begin VB.Menu calibmenu 
         Caption         =   "Chamber Volume"
         Index           =   7
      End
      Begin VB.Menu calibmenu 
         Caption         =   "Pressure &Gauge"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu calibmenu 
         Caption         =   "Cali&brate Comp. Reg."
         Index           =   9
      End
      Begin VB.Menu loadbackuplohm 
         Caption         =   "Load backup Lohm table"
      End
      Begin VB.Menu calibrateMv2 
         Caption         =   "Calibrate MV#2"
      End
      Begin VB.Menu mnuCalcHumidityTable 
         Caption         =   "Calc Humidity Table"
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Help"
      Index           =   6
      Begin VB.Menu helpmenu 
         Caption         =   "&View Configuration"
         Index           =   0
      End
      Begin VB.Menu helpmenu 
         Caption         =   "&About  CAPWIN"
         Index           =   1
      End
      Begin VB.Menu helpmenu 
         Caption         =   "&Help"
         Index           =   2
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Debug"
      Index           =   7
      Begin VB.Menu debugmenu 
         Caption         =   "Readat"
         Index           =   0
      End
      Begin VB.Menu debugmenu 
         Caption         =   "Stability debug"
         Index           =   1
      End
      Begin VB.Menu debugmenu 
         Caption         =   "DebugH20perm"
         Index           =   2
      End
      Begin VB.Menu debugmenu 
         Caption         =   "DebugBP"
         Index           =   3
      End
      Begin VB.Menu debugmenu 
         Caption         =   "LoopingDemo"
         Index           =   4
      End
      Begin VB.Menu commdebugmenu 
         Caption         =   "Debug comm port"
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Network"
      Index           =   8
      Begin VB.Menu networkmenu 
         Caption         =   "Setup"
      End
   End
End
Attribute VB_Name = "TitleScrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z

Dim ts$(66)                 ' Text strings for this form

Private Sub autoTestCMD_Click()
    runmenu_Click 1
End Sub

Private Sub AuxComm_OnComm()
    If AuxComm.InBufferCount > 0 Then abort_wait = True
End Sub

Private Sub calibmenu_Click(Index As Integer)

    Dim temp$, msg$, num_chambers_selected As Integer, i As Integer
    Dim mytemp As Single
    Dim compression_special As Boolean
    Dim T As Single

    checkReserveTankLevel

    Screen.MousePointer = 11
    Me.MousePointer = 11
    mainmenu(1).Enabled = False
    mainmenu(5).Enabled = False
    Static Indexx
    Indexx = Index

    ' ***** Initialize comm port ********
    Init_For_Ctrl True
    Zero_Reg
    'If ExtraPG Then
    'If Index = 3 Then
    Move_Valve 10, "C"
    Pres% = 1
    'Else
    ' Move_Valve 10, "O"
    ' Pres% = 2
    'End If
    'End If
    Move_Valve 0, "C"
    Move_Valve 2, "C"
    Dry_Chamber_Control "O"
    If slurry_tube_exists Then '6.71.123.19
        Move_Valve Slurry_tube_vent_valve, "C" 'slurry tube vent valve V29
    End If
    If H2OPERM Then
        Move_Valve 8, "C"
    End If
    Move_Valve 9, "C"

    ' they don't need to close the sample chamber for regulator calibration
    ' if they have liquid permeability since we can just close valve 4
    ' and we don't need to put the non-porous disk in place any more
    If Index = 3 Or Index = 4 Or (Index = 6 And H2OPERM = False) Then
        If autocompress Then
            ' run the compression presssure input here since they can
            ' clamp the chamber at this point
            Got_Value = -9 ' default so if they cancel we will know it
            GetCompressionPressure.Show 1
            If Got_Value <> -9 Then
                compression_pressure = Got_Value
                If compression_pressure = 0 And Not allowZeroCompression Then
                    compression_pressure = 20
                    compression_special = True
                Else
                    compression_special = False
                End If
                If creg_table_size% > 0 Then
                    If creg_table_pres!(creg_table_size%) < compression_pressure Then
                        compression_pressure = creg_table_pres!(creg_table_size%)
                    End If
                End If
                move_compression_regulator_to_pressure (compression_pressure)
                If compression_special Then compression_pressure = 0
            End If
        End If
    End If
    If Index = 2 Or Index = 3 Then
        temp$ = ts$(1)              ' "Please open and empty the sample chamber(s)."
        If multiChamberSystem And (manualMultiChamber = False) Then
            If Index = 2 Then
                ' for flow parameters, open everything
               '  temp$ = temp$ + "s" ' chambers is always plural, but not all languages form the plural this way.
                For current_unit% = 1 To chambers
                    Move_Valve -current_unit%, "O"
                Next current_unit%
                current_unit% = 1
            Else
                ' for lohm calibration, first count up the number of chambers
                ' while we are doing this, we can close all chambers
                num_chambers_selected = 0
                For i = 1 To chambers
                    Move_Valve -i, "C"
                    If selchamber(i) Then
                        num_chambers_selected = num_chambers_selected + 1
                    End If
                Next i
                If num_chambers_selected = 0 Then
                    MsgBox ts$(2)           ' "You don't have any chambers selected - Lohm calibration aborted"
                    GoTo end_of_sub
                End If
               ' If num_chambers_selected > 1 Then
                   ' temp$ = temp$ + "s" ' make plural
               ' End If
                i = num_chambers_selected ' number of chambers we have left
                current_unit% = 1
                While i > 0
                    ' find number of next chamber to at to list
                    While Not selchamber(current_unit%)
                        current_unit% = current_unit% + 1
                    Wend
                    temp$ = temp$ + str$(current_unit%)
                    current_unit% = current_unit% + 1
                    i = i - 1
                    If i > 1 Then
                        temp$ = temp$ + "," ' more than one more chamber to add
                    ElseIf i = 1 Then
                        temp$ = temp$ + " " + ts$(3)    ' "and" ' only one more chamber to add
                    End If
                Wend
                current_unit% = 1
            End If
        End If
        temp$ = temp$ + "." + vbCrLf + vbCrLf
        If Not itester Then
         temp$ = temp$ + ts$(4) + ", " + vbCrLf         ' "If the inlet pressure has changed"
         temp$ = temp$ + ts$(5) + vbCrLf                ' "the Flow Parameters Calibration must be run before" + vbCrLf
         temp$ = temp$ + ts$(6) + vbCrLf                ' "the Lohm & Max Air Flow Calibration is run"
        End If
        reply% = MsgBox(temp$, 65, ts$(7))              ' "Calibrations"
        If reply% = 2 Then
            GoTo end_of_sub
        End If
    End If

    If Index = 6 Then
        Move_Valve 2, "O" ' open the venting valve to start
        If H2OPERM Then
            ' don't open 4 to divert air just yet
            'Move_Valve 3, "O" ' divert air away from air chamber
            ' actually, close 4 to let air vent out while we zero everything
            Dry_Chamber_Control "O"
        ElseIf multiChamberSystem = False Or manualMultiChamber = True Then
            temp$ = ts$(8) + vbCrLf                     ' "Please insert the non-porous disk and seal the sample chamber."
            reply% = MsgBox(temp$, 65, ts$(7))          ' "Calibrations"
            If reply% = 2 Then
                GoTo end_of_sub
            End If
        Else
            Rem for multi-chamber machines, simply close the sample chamber valves
            For current_unit% = 1 To chambers
                Move_Valve -current_unit%, "C"
            Next current_unit%
            current_unit% = 1
        End If
    End If

    Msgform.lb.Visible = False
    Msgform.cancel.Visible = False
    ' don't need to compress at all if we are running regulator calibration and we
    '  are a liquid permeameter so we have valve 4
    ' don't need to compress at this time if we are running liquid flow test and
    '  we are a recirculation system, because they will be able to compress
    '  later when they do the recirculation setup
    If Index = 3 Or (Index = 4 And recirculation = False) Or (Index = 6 And H2OPERM = False) Then
        ' 6.71.47 use piston if it is there and they have set a compression pressure
        If autopiston Or (autocompress And compression_pressure <> 0 And current_unit% = 1) Then
            If safetydown Then
                If safetydowndoor Then
                    ' do auto door switch thing instead of key press thing
                    check_safety_door True
                Else
                    safetykeypress.mainlabel.Caption = ts$(62) ' "Piston About To Lower"
                    safetykeypress.Show 1
                End If
                If safety_canceled Then
                    GoTo end_of_sub
                End If
            ' autopiston is not allowed to ignore the safety down
            ElseIf autopiston Then
                check_safety_door True
                If safety_canceled Then
                    GoTo end_of_sub
                End If
            End If
            move_piston "O"
            '6.71.61 Give piston time to close
            'T = Timer
            'While (Timer - T <= pistonDelayTime): DoEvents: Wend
'            waitseconds pistonDelayTime
            '****
        End If
    End If
    Select Case Indexx
        Case 1
            Msgform.Label.Caption = ts$(9) + "..."          ' "Setting Valve Limits"
            Msgform.ok.Visible = False
            Msgform.Label1.Visible = False
            Msgform.Show 0
            Msgform.Refresh
            SetV2Limits
            msg$ = ts$(10)          ' "Calibration of Valve Limits is now automatic.  You no longer need to manually run the valve limits calibration"
        Case 2
            Msgform.Label.Caption = ts$(12) + "..."         ' "Calibrating Flow Meters"
            Msgform.ok.Visible = False
            Msgform.Label1.Visible = False
            Msgform.Show 0
            Msgform.Refresh
            SetLowFlow
            msg$ = ts$(13)              ' "Calibration of Flow Meters Complete"
        Case 3
            'Moved the following 11 lines of code from Run_Lohm_Table_Calibration to here
            'so it would only prompt once on multiple chamber machines
            MsgBox "Enter a filename to save to a location other than the default."
            
            'Prompt user for place and name to save group of lohm files
            Msgform.Label1.AutoSize = True
            Msgform.Label1.Caption = ts$(65)     ' "Lohm Calibration Completed"/"Pick a place to save this new lohm table."
            
            fsel_name$ = ""
            fsel_title$ = ts$(66)          ' "Save new lohm table"     ' ts$(42)                       ' "Backup Lohm Calibration File."
            fsel_path$ = EXE_Path$ + "parms\*.cal"
            fsel_io = False
            fsel TitleScrn.hwnd
            
            current_lohm_path$ = fsel_return$
            
            ' Write the current lohm table to user prefs
            WPPS Curr_U$, "lohmpath", current_lohm_path$, IFile$
            
            ' Update the main screen
            TitleScrn.Label1.Caption = get_filename_from_path(current_lohm_path$)
            TitleScrn.Label1.ToolTipText = current_lohm_path$
            
            Msgform.Label.Caption = ts$(14)         ' "Calculating Lohm"
            Msgform.StopTest.Visible = True
            Msgform.StopTest.Caption = ts$(64)  '               ' "End Calibration"
            Msgform.Label1.Visible = False
            Msgform.ok.Visible = False
            Msgform.Show 0
            Msgform.Refresh
            If (multiChamberSystem = False Or manualMultiChamber = True) And Not cartridge_tester Then
                Run_Lohm_Table_Calibration
            ElseIf cartridge_tester Then 'AJB 11-04-09
                'display pop asking user to select media or cartridge
                cartridgeLohmSelect.Show 1
                Run_Lohm_Table_Calibration
            Else
                For current_unit% = 1 To chambers
                    Msgform.Label.Caption = ts$(14) + " - Chamber " + Trim$(str$(current_unit%)) + "..."
                    If selchamber(current_unit%) Then
                        Move_Valve -current_unit%, "O"
                        Run_Lohm_Table_Calibration
                        Move_Valve -current_unit%, "C"
                    End If
                    ' if they stopped the calibration, don't do any more chambers
                    If Msgform.StopTest.Visible = False Then Exit For
                Next current_unit%
                current_unit% = 1
            End If
            msg$ = ts$(16)          ' "Calibration of Lohm Value Complete"
        Case 4
            Calc_MaxLQFlow
            msg$ = ts$(17)          ' "Calibration of Max Liquid Flow Complete"
        Case 5
            calc_diff_vol False
            msg$ = ts$(18)          ' "Calibration of Diffusion Volume Complete"
        Case 6
            calibrationComplete = False
            compregcal = False
            Screen.MousePointer = 0
            Me.MousePointer = 0
            calib_reg.Show 0
            While Not calibrationComplete
                DoEvents
            Wend
            msg$ = ts$(19)          ' "Calibration of Regulator Complete"
            If H2OPERM Then
                Dry_Chamber_Control "O"
            End If
        Case 7
            ' Calculation of chamber volume
            calc_diff_vol True
        Case 8
            Screen.MousePointer = 0
            Me.MousePointer = 0
            PGCalibrate.Show 0
            If GaugeCali_Aborted = True Then
               GaugeCali_Aborted = False
               Unload PGCalibrate
               msg$ = ts$(21)       ' "Calibration of Pressure Gauges is aborted"
            Else
               Unload PGCalibrate
               msg$ = ts$(22)       ' "Calibration of Pressure Gauges Complete"
            End If
        Case 9
            calibrationComplete = False
            compregcal = True
            Screen.MousePointer = 0
            Me.MousePointer = 0
            calib_reg.Show 1
            While Not calibrationComplete
                DoEvents
            Wend
            msg$ = ts$(23)          ' "Calibration of Compression Regulator Complete"
            compregcal = False
    End Select
    Unload Msgform
    MsgBox msg$, 0, ts$(24)         ' "Calibration Done"
end_of_sub:
    If autocompress Or autopiston Then
        If safetyup And Vpos(15) = 1 Then
            If safetyupdoor Then
                ' do auto door switch thing instead of key press thing
                check_safety_door False
            Else
                safetykeypress.mainlabel.Caption = ts$(63)     ' "Piston About To Raise"
                safetykeypress.cancelbutton.Visible = False
                safetykeypress.Show 1, progress
            End If
        End If
        move_piston "C"
    End If
    mainmenu(1).Enabled = True
    mainmenu(5).Enabled = True
    Screen.MousePointer = 0
    Me.MousePointer = 0
    
End Sub

Private Sub calibrateMv2_Click()
    ' ***** Initialize comm port ********
    Init_For_Ctrl True
    
    'initialize machine
        Zero_Reg
        
        zero_creg
        
        Move_Valve 15, "C"
        
        motorValveIndex = 0
        close_v2_completely
        
        motorValveIndex = 1
        close_v2_completely
        
        
    'display form
        Dim desiredFlow As Single
        Dim mv2Flow As Single
        Dim mode As Integer
        Dim time As Single
        Dim t0 As Single
        Dim startingFlow As Single
        
        ReadXReturnX4 1
        startingFlow = x5
        
        desiredFlow = FY2(1, 2)
        mv2Flow = 0
        mode = 0
        
        status.Label1.Caption = "Calibrating MV#2"
        status.Label2.Caption = "Desired Flow Rate: " + getFormattedFlow(desiredFlow)
        status.Label3.Caption = ""
        status.Command1.Visible = False
        status.Show
        
        inc_reg 100
        
        While mv2Flow < desiredFlow
            ReadXReturnX4 1
            mv2Flow = x5
            
            If mode = 0 Then
                'pulse mv2 to get flow
                Pulse_V2 0
                                                
                'update v2pos
                ReadXReturnX4 3
                V2POS = x4
                'look for flow
                t0 = Timer
                
                mv2Flow = 0
                time = 0
                
                While time < 5 And mv2Flow < (startingFlow + (desiredFlow * 0.25))
                    time = Timer - t0
                    ReadXReturnX4 1
                    mv2Flow = x5
                    status.Label3.Caption = "Current Flow: " + getFormattedFlow(mv2Flow)
                Wend
                
                If mv2Flow > (startingFlow + (desiredFlow * 0.25)) Then
                    mode = 1
                End If
                mv2Flow = 0
            ElseIf mode = 1 Then
                'we've gotten the motor valve cracked open, now we increment the regulator to get the
                'right flow rate.
                
                'Read current flow rate
                ReadXReturnX4 1
                mv2Flow = x5
                
                'determine how much to increment the regulator
                If mv2Flow < (desiredFlow * 0.5) Then
                    inc_reg 20
                ElseIf mv2Flow < (desiredFlow * 0.75) Then
                    inc_reg 10
                ElseIf mv2Flow < (desiredFlow * 0.9) Then
                    inc_reg 5
                Else
                    inc_reg 1
                End If
                
                
                'wait for the flow to settle out
                waitseconds 1
                
                Dim flowSet As Boolean
                flowSet = False
                Dim flow1 As Single
                Dim flow2 As Single
                
                While Not flowSet
                    flow1 = 0
                    flow2 = 0
                    Dim X As Integer
                    For X = 1 To 10
                        'read flow
                        ReadXReturnX4 1
                        flow1 = flow1 + x5
                    
                        'read flow again
                        ReadXReturnX4 1
                        flow2 = flow2 + x5
                        status.Label3.Caption = "Current Flow: " + getFormattedFlow(x5)
                    Next X
                    
                    flow1 = flow1 / 10#
                    flow2 = flow2 / 10#
                    'determine if it settled
                    If (flow2 - flow1) < (desiredFlow * 0.01) Then
                        flowSet = True
                    End If
                Wend
                'read the flow to see if we are high enough
                ReadXReturnX4 1
                mv2Flow = x5
            End If
            
        Wend
        
        Debug.Print "MV2 pos: " + str(V2POS)
        Debug.Print "REg pos: " + str(REGPOS)
        WPPS "Capstuff", "mv#2_start_pos", str$(V2POS), CSFile$
        WPPS "Capstuff", "mv#2_reg_pos", str$(REGPOS), CSFile$
        Zero_Reg
        status.Hide
End Sub

'Private Sub check_external_temperature_Click()
'
'    Dim s As Single
'
'    watlow_last_temperature = -99999        ' illegal value to detect error
'    s = read_external_watlow
'    If s = -99999 Then
'        MsgBox ts$(60)                      ' "Error: No communication with the external temperature controller"
'    Else
'        MsgBox ts$(61) + " =" + Str$(s)     ' "External temperature"
'    End If
'
'End Sub

Private Sub chselect_Click()
    selchambers.Show 1
End Sub

Private Sub Command1_Click()
    Dim s As Single

    GetaFile Me.hwnd
    If fsel_return$ = "" Then Exit Sub
    FResult = ""
    s = CalcGP(fsel_return$)
    MsgBox FResult + vbCrLf + "or " + str$(s)

End Sub

Private Sub Command2_Click()

    GetaFile Me.hwnd
    If fsel_return$ = "" Then Exit Sub
    FResult = ""
    CalcMFP fsel_return$
    MsgBox FResult

End Sub

Private Sub Command3_Click()
safetykeypress.Show 1
End Sub

Private Sub commdebugmenu_Click()

    Init_For_Ctrl True
    debugcomm.Show 1

End Sub

Private Sub convertData_Click(Index As Integer)
    MsgBox ("Please Select a file to convert ',' to '.' and '.' to ','")
    CommonDialog1.DialogTitle = "Open .CFT file"
    CommonDialog1.Filter = "CAPWIN Data file (*.cft)"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    Dim sFileText As String
    Dim iFileNo As Integer
    Dim oFileNo As Integer
    Dim test() As String
    Dim fixedLine As String
    Dim i As Integer, j As Integer
    Dim startConverting As Integer
    Dim newFileName() As String
    Dim currentLineNum As Integer
    Dim splitFile() As String
    Dim fullfile As String
    iFileNo = FreeFile
    Open CommonDialog1.filename For Input As iFileNo
    oFileNo = FreeFile
    newFileName = split(CommonDialog1.filename, ".cft", , vbTextCompare)
    Open newFileName(0) + "-converted.cft" For Output As oFileNo
    startConverting = 9
    j = 0
    currentLineNum = 1
    Do While Not EOF(iFileNo)
        Line Input #iFileNo, sFileText
        If currentLineNum = 2 Then startConverting = 2 + val(sFileText)
        If currentLineNum > startConverting Then
            test = split(sFileText, ",")
            fixedLine = ""
            i = 0
            Do While i <= UBound(test)
                If Not (fixedLine = "") Then
                    fixedLine = fixedLine + "." + Replace(test(i), ".", ",")
                Else
                    fixedLine = Replace(test(i), ".", ",")
                End If
                i = i + 1
            Loop
            If (InStr(fixedLine, Chr(32)) = 1 And InStr(fixedLine, ",") = 2) Then fixedLine = "0" + LTrim$(fixedLine)
            Print #oFileNo, fixedLine
        Else
            Print #oFileNo, sFileText
        End If
        j = j + 1
        currentLineNum = currentLineNum + 1
    Loop
    Close #iFileNo
    Close #oFileNo
    MsgBox ("Done")
End Sub

Private Sub d1a_Click(Index As Integer)
BorderColor.Show 0

End Sub

Private Sub debugmenu_Click(Index As Integer)

    debugmenu(Index).Checked = Not debugmenu(Index).Checked
    
    Select Case Index
        Case 0  ' Readat
            readatenabled = debugmenu(Index).Checked
        Case 1  ' stability debug
            'stability_debug = debugmenu(Index).Checked
        Case 2  ' debug h2o perm
            debugH20Perm = debugmenu(Index).Checked
        Case 3  ' debug BP
            debugBP = debugmenu(Index).Checked
        Case 4  ' Looping demo
            LoopingDemo = debugmenu(Index).Checked
        Case Else
    End Select
    
End Sub

Private Sub executeReportCMD_Click()
    Reportmenu_Click 1
End Sub

Private Sub exitMenu_Click(Index As Integer)
    Unload Me
End Sub

Private Sub experimental2Label_Click()
    Select_Test_New.Show
End Sub

Private Sub filemenu_Click(Index As Integer)

    Dim def$, editgrpnm$, Ret$, i%, XTemp$, DFile$, temp$, T%
    Dim r As Long
    Dim Exists As Boolean
    Dim numUsers%, usernum%
    Dim LI As Integer
    Dim unit%, u$
    
    Select Case Index
        Case 1
            T_Select$ = "USER"
            Selection.Show 1
            If T_Select$ <> "Cancel" Then GoSub Set_Cur_Usr
        Case 2
            NewSID.Show 1
            GoSub Set_Cur_Usr
        Case 3
            T_Select$ = "DEL"
            Selection.Show 1
        Case 4
            If UCase$(Curr_U$) = "DEFAULT" Then
                MsgBox ts$(26), , ts$(27)     ' "You can not edit the DEFAULT group name.   Please switch to a different group or create a new group."/"Edit group name"
                Exit Sub
            End If
            def$ = Label3.Caption
            editgrpnm$ = UCase$(InputBox(ts$(28), ts$(27), Label3.Caption))   ' "Please enter the new group name."/"Edit Group Name"
            editgrpnm$ = LTrim$(RTrim$(editgrpnm$))
            If editgrpnm$ = UCase$(Curr_U$) Then
                editgrpnm$ = ""
            End If
            If editgrpnm$ <> "" Then
                'save new name
                usernum% = 0
                Exists = False
                Ret$ = String$(255, vbNullChar)
                r = GPPS("default", "numusers", "1", Ret$, 255, IFile$)
                numUsers% = val(Ret$)
                For i% = 1 To numUsers%
                    r = GPPS("default", "user" + Format$(i%), "", Ret$, 255, IFile$)
                    If UCase$(nulltrim(Ret$)) = UCase$(Curr_U$) Then
                        usernum% = i%
                    End If
                    If UCase$(nulltrim(Ret$)) = editgrpnm$ Then
                        MsgBox ts$(29)          ' "New name already exists - name change aborted"
                        Exists = True
                        Exit For
                    End If
                Next i%
                If usernum% = 0 Then
                    MsgBox ts$(30)              ' "Something is wrong - current user index couldn't be found"
                    Exists = True ' set to true to bypass changing name
                End If
                If Not Exists Then
                    XTemp$ = editgrpnm$
                    ' try renaming the user folder
                    On Error GoTo cant_rename_folder
                    Name EXE_Path$ + "users\" + Curr_U$ As EXE_Path$ + "users\" + XTemp$
renamed_folder_ok:
                    On Error GoTo 0
                    For unit% = 1 To chambers
                        If unit% = 1 Then u$ = "" Else u$ = Format$(unit%)
                        load_user_stuff u$
                    Next unit%
                    WPPS Curr_U$, vbNullString, vbNullString, IFile$
                    WPPS "default", "user", XTemp$, IFile$
                    WPPS "default", "user" + Format$(usernum%), XTemp$, IFile$
                    ' when renaming a user, if the parameters point to the old user directory,
                    ' they now need to point to the new user directory
                    For unit% = 1 To chambers
                        check_group_change Curr_U$, XTemp$, TPFDRY$(unit%)
                        check_group_change Curr_U$, XTemp$, TPFWET$(unit%)
                        check_group_change Curr_U$, XTemp$, OutFilename$(unit%)
                    Next unit%
                    check_group_change Curr_U$, XTemp$, path(0)
                    check_group_change Curr_U$, XTemp$, path(1)
                    Curr_U$ = XTemp$
                    
                    
                    For unit% = 1 To chambers
                        If unit% = 1 Then u$ = "" Else u$ = Format$(unit%)
                        save_user_stuff u$
                    Next unit%
                End If
done_new_name:
                GoSub Set_Cur_Usr

            End If
        Case 5, 6, 7, 8, 9, 10

     '   Case 15
     '       Unload Me
        Case Else
    End Select
    Exit Sub
bad_new_name:
    MsgBox ts$(32)          ' "Error changing name"
    Resume done_new_name
cant_rename_folder:
    ' maybe it already exists - try copying files
    If file_copy(XTemp$, "capgraph.ini") Or file_copy(XTemp$, "caprep.ini") Or file_copy(XTemp$, "rp_cover.txt") Or file_copy(XTemp$, "default.tpf") Then
        MsgBox ts$(33)      ' "Error creating new user folder"
        Resume done_new_name
    End If
    Resume renamed_folder_ok
Set_Cur_Usr:
    T_Select$ = ""
    WPPS "default", "user", Curr_U$, IFile$
    load_user_global_stuff
    Return
    
End Sub

Private Sub Form_Load()
    Debug.Print App.path 'AJB 10-21-09
    Dim filenum As Integer
    Dim accessLevel As Long
    Dim temp$, Ret$, T%, n%, i%
    Dim r As Long, freq As Long
    Dim xpath As String     'edc 12-07-06 leftside of EXE_Path
    Dim slash As Integer    'edc 12-07-06 location of the slash in EXE_Path
    Dim expath As String    'edc 12-07-06 whole path to go in the title bar
    Dim rawCrypt As String
    Dim errNum As Integer
    
    On Error GoTo UnknownErrorHandler
    errNum = 1
    
    ' setup comm error log array
    For i% = 1 To 100
        cel1(i) = 0
        cel2(i) = 0
        cel3(i) = 0
        cel_i = 1 ' where to put next one
    Next i%

    errNum = 2

    penetrometer_select = 1
    current_unit = 1
    piston_status = 0 ' unknown initial piston status
    just_did_vacuum_purge = False
    last_reg_status = False
    CFAnal = False
    GasPerm = False
    TestScreenVisible = False

    '''''''''''''''''''''''''''''''
    '8/31/12- Sam Bouabane + Sean Vesley
    'these variables hold the beginning parameters for pressure hold
    '''''''''''''''''''''''''''''''
    startThresholdMax = 500
    startThresholdMin = 0
    
    
    S_Version = App.Major & "." & App.Minor & "." & App.Revision '& "x"
    errNum = 3

    App.Title = App.EXEName
    ' before setting the path based on the application executable,
    ' first check the command line for an override
    EXE_Path$ = Command$
    errNum = 4
    
    If Len(EXE_Path$) > 0 Then
        If Dir$(EXE_Path$ + ProgAbbv + "win.ini") = "" Then
            EXE_Path$ = App.path
        End If
    Else
        EXE_Path$ = App.path
    End If
    If Right$(EXE_Path$, 1) <> "\" Then EXE_Path$ = EXE_Path$ + "\"
    'MsgBox Exe_Path$

    errNum = 5
    
    On Error Resume Next
    filenum = FreeFile
    Open EXE_Path$ + ProgAbbv + "win.ini" For Input As #filenum
    If Err <> 0 Then
        temp$ = EXE_Path$
        EXE_Path$ = "c:\program files\" + ProgAbbv + "win\"
        Close #filenum
        Err = 0
        filenum = FreeFile
        Open EXE_Path$ + ProgAbbv + "win.ini" For Input As #filenum
        If Err <> 0 Then
            If Mid$(temp$, 4, 5) = "CODE\" Then EXE_Path$ = Left$(temp$, 3) + "wins\" + ProgAbbv + "win\"
            Close #filenum
            Err = 0
            filenum = FreeFile
            Open EXE_Path$ + ProgAbbv + "win.ini" For Input As #filenum
            If Err <> 0 Then
                Close #filenum
                filenum = 0
                Err = 0
                'use fsel to locate file.  If fsel_return comes back as a "", end prog
                'MsgBox ts$(36) + " " & ProgAbbv & "Win.ini " + ts$(37)   ' "Unable to locate"/"file.  Please locate it."
                MsgBox "Unable to locate " & ProgAbbv & "Win.ini file. Please locate it."
                fsel_name$ = LCase$(ProgAbbv) + "win.ini"
                fsel_title$ = "Please Find the Initializaion File"      'ts$(38)       '
                fsel_path$ = App.path + "\*.ini"
                fsel_io = True
                fsel Me.hwnd
                If fsel_return$ = "" Then
                    MsgBox "Unable to start CapWin without 'capwin.ini' file.  Check the Windows Explorer to locate the file."  'ts$(39)          '
                    End
                End If
                EXE_Path$ = Left$(fsel_return$, Len(fsel_return$) - 10)
                'Do Until Right$(exe_path$, 1) = "\"
                '    exe_path$ = Left$(exe_path$, Len(exe_path$) - 1)
                'Loop
            End If
        End If
    End If
    
    errNum = 6
    
    On Error GoTo UnknownErrorHandler
    If filenum > 0 Then Close #filenum
    filenum = 0
    
    SetupLanguage EXE_Path$, "CapWin"
    'LoadTextStrings

    errNum = 7

    If App.PrevInstance Then
        MsgBox App.EXEName + " " + ts$(34), 0, ts$(35)     ' "is already running in another window."/"Program Active"
        Unload Me
        End
    End If
    
    errNum = 8
    
    
    Set UAC_Crypt = New vbCrypt.EncryptionTools
    

    UACFile$ = EXE_Path$ + "uac.ini"
    UAC_Enabled = IIf(gpps2("UAC", "UAC_Enabled", UACFile, "N") = "Y", True, False)
    UAC_UserCount = val(gpps2("UAC", "UAC_UserCount", UACFile, "0"))
    UAC_CurrentUser = val(gpps2("UAC", "UAC_CurrentUser", UACFile, "1"))
    If UAC_UserCount > 0 Then
        accessLevel = val(gpps2("UAC", "UAC_AL", UACFile, "0"))
        ReDim UAC_Users(UAC_UserCount)
        For i% = 1 To UAC_UserCount
            UAC_Users(i%).Username = gpps2("UAC", "UAC_Username" + Trim$(str$(i%)), UACFile, "")
            rawCrypt = gpps2("UAC", "UAC_Password" + Trim$(str$(i%)), UACFile, "")
            UAC_Users(i%).Password = UAC_Crypt.URLDecodeBinaryData(rawCrypt)
            UAC_Users(i%).accessLevel = (accessLevel And (2 ^ (i% - 1)))
        Next i%
        If UAC_Enabled Then
            frmLogin.Show vbModal
            If Not UAC_LoggedIn Then
                Unload Me
                End
            End If
            lblUsername.Visible = True
            lblUsername.Caption = "Current User: " + UAC_Users(UAC_CurrentUser).Username
            mnuLogout.Visible = True
            mnuLogoutSepLine.Visible = True
            mode.Visible = False
            filemenu(10).Visible = False
        End If
    End If
    
    errNum = 9
    
    capcommon.LoadTextStrings
    
    Decimal_Point$ = Mid$(Format$(1.2, "#.#"), 2, 1)
    ExtraPG = False
    HelpFile$ = EXE_Path$ + "CAPHELP.HLP"
    IFile$ = EXE_Path$ + "CapUsers.ini"
    CSFile$ = EXE_Path$ + "capwin.ini"
    currentLanguagePath$ = gpps2("default", "current_language_path", IFile, "CapWinLanguageEN.ini")
    language$ = App.path & "\languages\" & currentLanguagePath
    CAPFLOW.LoadTextStrings 'edc 03-05-07 changed the location so as to set the language first
    LoadTextStrings
    'edc 03-08-07 added
    'MinimumReliableHighFlowRate = gpps2("Capstuff", "MinimumReliableHighFlowRate", CSFile$, 0)
    ' check for gasflowconversion.ini file
    errNum = 10

    Ret$ = String$(255, " ")
    r = GPPS("main", "gas1", "", Ret$, 255, EXE_Path$ + "gasflowconversion.ini")
    If r = 0 Then
        ' create default file
        WPPS "main", "gas1", "Nitrogen", EXE_Path$ + "gasflowconversion.ini"
        WPPS "main", "conversion1", "1", EXE_Path$ + "gasflowconversion.ini"
        WPPS "main", "current", "1", EXE_Path$ + "gasflowconversion.ini"
    End If
    
    ' check to see if there is more than one gas.  If not, we don't use conversion feature
    errNum = 11
    
    r = GPPS("main", "gas2", "", Ret$, 255, EXE_Path$ + "gasflowconversion.ini")
    If r = 0 Then
        Me.GasFrame.Visible = False
        Me.GasLabel.Visible = False
        'Me.Shape1.Height = 6705
        'Me.Frame1.Height = 7205     'edc 12-07-06 edit for border Color
        'Me.Height = 8115            'edc 12-07-06 edit for border color
        gasflowconversionfactor = 1
    Else
        GPPS "main", "current", "1", Ret$, 255, EXE_Path$ + "gasflowconversion.ini"
        i% = val(Ret$)
        If i% < 1 Then i% = 1
        GPPS "main", "gas" + Format$(i%), "Air", Ret$, 255, EXE_Path$ + "gasflowconversion.ini"
        Me.GasLabel.Caption = nulltrim(Ret$)
        GPPS "main", "conversion" + Format$(i%), "1", Ret$, 255, EXE_Path$ + "gasflowconversion.ini"
        gasflowconversionfactor = val(Ret$)
        If gasflowconversionfactor <= 0 Then gasflowconversionfactor = 1
    End If
    
    errNum = 12
    
    Rem turn off normally hidden menu items
    Rem we leave these visible by default so that we can
    Rem see them in the form designer
    calibmenu(1).Visible = False
    calibmenu(4).Visible = False
    calibmenu(5).Visible = False
    calibmenu(6).Visible = False
    calibmenu(9).Visible = False
    Rem the machines that use these features will make
    Rem them visible again, but this way the machines
    Rem that don't need them will not have to hide them.
    compregcal = False ' just in case
    
    PWFACTR = 1
    
    'All References to Users after this actually refer to Groups
    
    Rem first test the user file to make sure that the user list
    Rem is stored in the default position
    errNum = 13
    
    Ret$ = String$(255, " ")
    r = GPPS("default", "numusers", "foo", Ret$, 255, IFile$)
    If nulltrim(Ret$) = "foo" Then
        Rem we need to create the user list for them
        Rem first, see if there is a user list in capwin.ini
        r = GPPS("userlist", "numusers", "foo", Ret$, 255, CSFile$)
        n% = val(Ret$)
        If nulltrim(Ret$) = "foo" Or n% < 1 Then
            Rem no user list in the capwin.ini file
            Rem create one
            WPPS "default", "numusers", "1", IFile$
            WPPS "default", "user", "DEFAULT", IFile$
            WPPS "default", "user1", "DEFAULT", IFile$
        Else
            Rem copy the user list from capwin.ini
            WPPS "default", "numusers", str$(n%), IFile$
            r = GPPS("userlist", "user", "default", Ret$, 255, CSFile$)
            WPPS "default", "user", nulltrim(Ret$), IFile$
            For i% = 1 To n%
                r = GPPS("userlist", "user" + Format$(i%), "default", Ret$, 255, CSFile$)
                WPPS "default", "user" + Format$(i%), nulltrim(Ret$), IFile$
            Next i%
        End If
    End If
    r = GPPS("default", "user", "default", Ret$, 255, IFile$)
    Curr_U$ = UCase$(nulltrim(Ret$))
    
    errNum = 14
    
    CheckUserGroups
    
    Get_First = False
    Init_For_Ctrl False ' don't do any actual talking with the machine
    switchingMVs = False
    Rem unitnumber should now be set
'   unitnumber has already updated the caption in init_for_ctrl
'    If unitnumber <> 0 Then
'        Me.Caption = Me.Caption + " - Unit" + Str$(unitnumber)
'    End If
    errNum = 15

    If (capwin_feature_number And 64) = 64 Then
        ' Has a Solenoid Valve instead of a Motor Valve
        calibrateMv2.Visible = False
    Else
        calibrateMv2.Visible = True
    End If

    Me.runmenu(6).Visible = (autocompress Or autopiston) And _
       (safetyup = False Or safetyupdoor = True) And _
       (safetydown = False Or safetydowndoor = True)
       
    runmenu(8).Visible = (reservoirTemperature <> 0) Or _
        (airTemperature <> 0) Or _
        (cabinetTemperature <> 0) Or _
        (bubblerTemperature <> 0) Or _
        (dryChamberTemperature <> 0) Or _
        (wetChamberTemperature <> 0) Or _
        (hydroHeadTemperature <> 0) Or _
        (mullenTemperature <> 0)
    
    mainmenu(8).Visible = network_connection_enabled
    
    errNum = 16
    
    load_user_global_stuff

    errNum = 17
    
    loadMicrometers
    readAllMicrometers
    If numberOfMicrometers > 1 Then
        mnuMicrometerSelect.Visible = True
        mnuMicrometer(2).Visible = IIf(numberOfMicrometers > 2, True, False)
        currentMicrometer = val(gpps2("Capstuff", "CurrentMicrometer", CSFile$, str$(0)))
        For i = 0 To numberOfMicrometers - 1
            mnuMicrometer(i).Checked = False
        Next i
        mnuMicrometer(currentMicrometer).Checked = True
    Else
        mnuMicrometerSelect.Visible = False
    End If
    
    errNum = 170

    ' Test to see if hardware supports using performance counter for timing
    freq = QueryPerformanceFrequency(PerformanceFrequencyValue)
    If freq = 0 Then
        MsgBox ts$(59)      ' "Error:  Your hardware is not compatible with this software - performance counter hardware not detected"
        End
    End If
    'edc 12-06-07
    'The folowing lines add border color and the executable pathe to the
    'title car of the form as well as set the SubCaption for the rest of
    'the rest of the forms
    Me.BackColor = lngBorderColor
     xpath = Left(EXE_Path$, Len(EXE_Path) - 1)
      slash = InStrRev(xpath, "\")
    expath = Left(EXE_Path$, 3) & "..." & Right(xpath, Len(xpath) - slash)
    SubCaption = expath & "\" & App.EXEName & ".exe"
    Me.Caption = Me.Caption & "    " & SubCaption       '12-07-06 edc
    If simpleqc_enable Then
        Me.WindowState = 1
    End If
    errNum = 18
    
    Me.Show
    Refresh
    temperature_array_size = 10
    ReDim temperature1(10), temperature2(10)
    On Error GoTo bad_capwinrg
    filenum = FreeFile
    Open EXE_Path$ + "capwinrg.cal" For Input As #filenum
    Input #filenum, n%
    reg_table_size%(0) = n%
    ReDim reg_table_pos(n%) As Long, reg_table_pres!(n%)
    For i% = 1 To n%
        DoEvents
        Input #filenum, reg_table_pos(i%), reg_table_pres!(i%)
    Next i%
    
    ' Check visibility and status of the debug menu
    If debugMenuVisible Then
        mainmenu(7).Visible = True
        debugmenu(0).Checked = readatenabled        ' "READAT"
        debugmenu(1).Checked = stability_debug      ' "stability debug"
        debugmenu(2).Checked = debugH20Perm         ' "liq.perm debug"
        debugmenu(3).Checked = debugBP              ' bubble point debug
        debugmenu(4).Checked = LoopingDemo          ' Running a looping BP/GP test
        ' Next item is debug comm port, which is a click-and-play setup
    Else
        mainmenu(7).Visible = False
    End If
    
    ' 6.71.60 Check to see if "experimental" warning should be visible (only if version number ends in "x")
    experimentalLabel.Visible = (LCase$(Right$(S_Version, 1)) = "x")
    experimental2Label.Visible = (LCase$(Right$(S_Version, 1)) = "x")
'    If hasHumidityControls Then
'        loadHumidityTable
'    Else
        mnuCalcHumidityTable.Visible = False
'    End If

    zeroRegVentTime = 30000 'set vent time on reg to 30 seconds, only applies to machines with a bubbler for humidity
    
    
done_capwinrg:
    Close #filenum
    If dualregulator Then
        On Error GoTo bad_capwinrg2
        filenum = FreeFile
        Open EXE_Path$ + "capwinrg2.cal" For Input As #filenum
        Input #filenum, n%
        reg_table_size%(1) = n%
        ReDim reg_table_pos2(n%) As Long, reg_table_pres2!(n%)
        For i% = 1 To n%
            DoEvents
            Input #filenum, reg_table_pos2(i%), reg_table_pres2!(i%)
        Next i%
done_capwinrg2:
        Close #filenum
    End If
    On Error GoTo bad_capwincrg
    filenum = FreeFile
    Open EXE_Path$ + "capwincompreg.cal" For Input As #filenum
    Input #filenum, n%
    creg_table_size% = n%
    ReDim creg_table_pos(n%) As Long, creg_table_pres!(n%)
    For i% = 1 To n%
        DoEvents
        Input #filenum, creg_table_pos(i%), creg_table_pres!(i%)
    Next i%
done_capwincrg:
    Close #filenum
    errNum = 19
    On Error GoTo UnknownErrorHandler
    Screen.MousePointer = 0
    If simpleqc_enable Then
        Me.Timer1.Enabled = True
    End If
    If liqpermonly Then
        calibmenu(7).Visible = False
    End If
    Exit Sub
bad_capwinrg:
    reg_table_size%(0) = 0
    ' dim table anyway to avoid errors
    ReDim reg_table_pos(1) As Long, reg_table_pres!(1)
    Resume done_capwinrg
bad_capwinrg2:
    reg_table_size%(1) = 0
    ' dim table anyway to avoid errors
    ReDim reg_table_pos2(1) As Long, reg_table_pres2!(1)
    Resume done_capwinrg2
bad_capwincrg:
    creg_table_size% = 0
    ' dim table anyway to avoid errors
    ReDim creg_table_pos(1) As Long, creg_table_pres!(1)
    Resume done_capwincrg
    
    leak_test_delay = -1
    
    readAllFlows = False
    Exit Sub
UnknownErrorHandler:
    If errNum = 8 Then
    MsgBox "Error " + str$(errNum) + "VB Crypt is not installed, Please Install vbCrypt.dll into your System32 folder."
    Else
    MsgBox "Unknown Error " + str$(errNum) + vbCrLf + str$(Err.Number) + " " + Err.Description, vbInformation
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)

    Dim numdata&
    Dim r As Long
    
    If simpleqc_enable Then Exit Sub
    numdata& = 0
    On Error Resume Next
    r = WinHelp(hwnd, HelpFile$, Help_Quit, ByVal numdata&)
    reply% = MsgBox(ts$(40), 17, "CapWin")      ' "Exit CapWin and Abort any Running Tests ?"
    If reply% = 1 Then
        If log_comm Then
            log_append "all done"
        End If
        closeAllMicrometers
        End
    Else
        cancel = -1
        On Error GoTo 0 ' added by HDS on 6/4/99
    End If

End Sub


Private Sub GasLabel_Click()
' select a new gas
    T_Select$ = "GAS"
    Selection.Show 1
    If T_Select$ = "Cancel" Then Exit Sub
End Sub

Private Sub Helpmenu_Click(Index As Integer)

    Dim numdata&
    Dim Exists As Boolean
    Dim directory As String
    directory = EXE_Path$ + "\CapWinHelp"
     Exists = DirExists(directory)
    numdata& = 0
    Select Case Index
        Case 0
            Get_First = -1
                Get_Capstuff
            Get_First = 0
        Case 1
            AboutBox.Show 1
        Case 2
            chooseHelp (Exists)
        Case Else
    End Select
    
End Sub
Private Sub chooseHelp(Exists As Boolean)
Dim r As Long
Dim ShellCode As Long
Dim numdata&
numdata& = 0
'The following code checks to see if the new help system has been
            'added. If it has the new system is called and displayed in the users
             'default browser, if not the old help system is utilized. ecjc 6-14-05
            If Exists = True Then
                ShellCode = ShellExecute(Me.hwnd, "open", EXE_Path$ + "\CapWinHelp\index.htm", "", "", 1)
            Else
            r = WinHelp(hwnd, HelpFile$, Help_Index, ByVal numdata&)
            End If
End Sub
'This function is designed to determin weather the CapWinHelp file exists ecjc 6-14-05.

Function DirExists(path As String) As Boolean
On Error Resume Next
DirExists = (Dir$(path & "\nul") <> "")
End Function
Private Sub loadbackuplohm_Click()

    fsel_name$ = ""
    fsel_title$ = ts$(41)                   ' "Backup Lohm Calibration File"
    fsel_path$ = EXE_Path$ + "parms\*.cal"
    fsel_io = True
    fsel TitleScrn.hwnd
    
    If fsel_return$ <> "" Then
    
        If InStr(fsel_return$, ".cal") <> InStrRev(fsel_return$, ".cal") Then
            current_lohm_path$ = Left(fsel_return$, InStr(fsel_return$, ".cal") + 3)
        Else
            current_lohm_path$ = fsel_return$
        End If
            
        TitleScrn.Label1.Caption = get_filename_from_path(current_lohm_path$)
        TitleScrn.Label1.ToolTipText = current_lohm_path$
        
        WPPS Curr_U$, "lohmpath", current_lohm_path$, IFile$
        
    End If
    
End Sub

Private Sub MainComm_OnComm()
    If MainComm.InBufferCount > 0 Then abort_wait = True
End Sub

Private Sub manualTestCMD_Click()
    runmenu_Click 2
End Sub

Private Sub mnuCalcHumidityTable_Click()
    
    'below commands are taken from the lohm calibration initialization code, they're put here because
    'the humidity table calibration is similar to lohm in what we need to do, namely increase the flow
    'to a series of points and measure something
    Init_For_Ctrl True
    Zero_Reg
    
    Move_Valve 10, "C"
    
    Move_Valve 0, "C"
    Move_Valve 2, "C"
    Dry_Chamber_Control "O"
    
    If slurry_tube_exists Then '6.71.123.19
        Move_Valve Slurry_tube_vent_valve, "C" 'slurry tube vent valve V29
    End If
    If H2OPERM Then
        Move_Valve 8, "C"
    End If
    Move_Valve 9, "C"
    
    'enable valve for humidity
    Move_Valve 3, "C"
    
    'Display status box
    Msgform.Label.Caption = "Calculating humidity table..."
    Msgform.StopTest.Visible = True
    Msgform.lb.Visible = False
    Msgform.ok.Visible = False
    Msgform.Show 0
    Msgform.Refresh

    calcHumidityTable
        
    MsgBox "Done calculating humidity table."

End Sub

Private Sub mnuLogout_Click()
    Me.Visible = False
    UAC_LoggedIn = False
    frmLogin.Show vbModal
    If Not UAC_LoggedIn Then
        Unload Me
        End
    End If
    update_super_mode
    Me.Visible = True
End Sub

Private Sub mnuMicrometer_Click(Index As Integer)
    Dim i As Integer
    
    currentMicrometer = Index
    For i = 0 To 2
        mnuMicrometer(i).Checked = False
    Next i
    mnuMicrometer(Index).Checked = True
    WPPS "Capstuff", "CurrentMicrometer", str(Index), CSFile$
End Sub

Private Sub modemenu_Click(Index As Integer)
' Handle selection of modes

    Select Case Index
        Case 0 ' enter supervisor mode if we are currently in user mode
            If Not supervisor Then
                If superpass$ <> "" Then
                    GetValue.Label1.Caption = ts$(31) + ":"  ' "Enter Password"
                    GetValue.Text1.Text = ""
                    GetValue.Text1.SelStart = 0
                    GetValue.Text1.SelLength = 0
                    GetValue.Label1.Tag = "text"
                    GetValue.Continue.default = True
                    GetValue.Show 1
                    GetValue.Label1.Tag = ""
                    If Got_Value = -9 Then Exit Sub
                    If Got_Text <> superpass$ Then Exit Sub
                End If
                supervisor = True
                update_super_mode
                Exit Sub
            End If
        Case 1 ' enter user mode if we are currently in supervisor mode
            If supervisor Then
                supervisor = False
                update_super_mode
            End If
        Case 2
            ' simple QC Mode turn on
            WPPS "Capstuff", "SimpleQC", "1", CSFile$
            simpleqc_enable = True
            setup_simple_qc_mode
        Case 4
            ' Change password
            GetValue.Label1.Caption = ts$(43) + ":"  ' "Enter New Password"
            GetValue.Text1.Text = ""
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = 0
            GetValue.Label1.Tag = "text"
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value = -9 Then Exit Sub
            superpass$ = Got_Text
            WPPS "Capstuff", "superpass", superpass$, CSFile$
    End Select
    
End Sub

Private Sub modifymenu_Click(Index As Integer)

    Select Case Index
        Case 0
            autoparm.Show 0
        Case 1
            Manfile.Show 0
        Case 2
            Curve_Ave = 1
            Curve.Show 0
        Case 3
            EditForm.Show 0
        Case 4
            Curve_Ave = 2
            Curve.Show 0
            
    End Select
    
End Sub

Private Sub networkmenu_Click()
network.Show
End Sub


Private Sub parametersCMD_Click()
    modifymenu_Click 0
End Sub

Private Sub Picture1_DblClick()

    Dim r As Double
    
    If Not supervisor Then Exit Sub
    r = Shell("notepad.exe " + CSFile$, 1)
   ' r = Shell(EXE_Path + "initializer.exe")

End Sub

Private Sub Picture4_DblClick()
    Dim r As Double
    
    If Not supervisor Then Exit Sub
    r = Shell("notepad.exe " + CSFile$, 1)
End Sub

Private Sub Preferences_Click()
    prefsForm.Show
End Sub

Private Sub prefs_Click(Index As Integer)
    prefsForm.Show
End Sub

Private Sub psl_Click(Index As Integer)
' Pressure step lists

    stepListForm.Show 0

End Sub

Private Sub Reportmenu_Click(Index As Integer)

    Dim tempx As Double

    Select Case Index
        Case 1
            tempx = Shell(EXE_Path$ + "Caprep.exe", 1)
        Case 2
            tempx = Shell("notepad " + EXE_Path$ + "users\" + Curr_U$ + "\rp_cover.txt", 1)
      
    End Select
    
End Sub

Private Sub runmenu_Click(Index As Integer)
    Dim compression_special As Boolean
    
    Me.MousePointer = 11
    mainmenu(1).Enabled = False
    mainmenu(5).Enabled = False
    Init_For_Ctrl True
    
    checkReserveTankLevel
    
    If Not cartridge_tester Then current_unit% = 1
    
    flow_status = 0 ' initial condition - we don't know what the flow status really is
    ' added index 4 so it uses compression for leak test 6.71.47
    If Index = 1 Or Index = 4 Or Index = 7 Then
        If autocompress Then
            ' run the compression presssure input here since they can
            ' clamp the chamber at this point
            Got_Value = -9 ' default so if they cancel we will know it
            GetCompressionPressure.Show 1
            If Got_Value <> -9 Then
                compression_pressure = Got_Value
                If compression_pressure = 0 And Not allowZeroCompression Then
                    compression_pressure = 20
                    compression_special = True
                Else
                    compression_special = False
                End If
                If (creg_table_size% > 0) And (creg_table_pres!(creg_table_size%) < compression_pressure) Then
                    compression_pressure = creg_table_pres!(creg_table_size%)
                End If
                move_compression_regulator_to_pressure (compression_pressure)
                If compression_special Then compression_pressure = 0
            End If
        End If
    End If
    Select Case Index
        Case 1
            If multiChamberSystem And (manualMultiChamber = False) Then
                Do Until selchamber(current_unit%)
                    current_unit% = current_unit% + 1
                    If current_unit% > chambers Then
                        MsgBox ts$(44)          ' ("Error:  No chambers selected for testing")
                        GoTo click_done
                    End If
                Loop
            End If
            RUNNING = False ' this will be set to true if we are really
                            ' running a test
            ' autocompress prompts moved to above
            If recirculation Then
                run_recirculation_setup 0 ' initial run of recirculation setup
                If flow_status = 0 Then
                    ' they canceled the test setup
                    GoTo click_done
                End If
            End If
            If ExtraPG Then Pres% = 2 Else Pres% = 1
            intest = True
            save_setup_data_flag = False
            first_test_setup = True
            Testscrn.Show 1
            
            checkReserveTankLevel
            
            If RUNNING Then
                checkSampleChamberDiverterValve
            End If
            
            'AJB this is the entry point before any test
            If testGasStatus And RUNNING And Not H2OPERM Then testGas
            
            'ajb 12-26-09
            If hasMultipleMVs Then monitorFlow = False
            
            'JF 02-10-2010
            'Add a function to set the temperatures and possible wait for them to achieve the
            'target temperature
            If useTemperatureControlForAuto(current_unit%) Then
                PrepareTemperatures
            End If
            
            'JF 02-16-2010
            'Add a function to set the humidity and wait for the system to achieve that
            'target humidity
            'If hasHumidityControls And RUNNING Then PrepareHumidity
            
            ' if save_setup_data_flag is true but RUNNING is false,
            ' these routines will exit after saving the information
            If sequentialTesting And save_setup_data_flag Then
                MsgBox ("RunSequentialTest test type")
                RunSequentialTest
            Else
                If save_setup_data_flag Then
                    If multiChamberSystem And manualMultiChamber = False Then
                        'MsgBox ("RunMultipleTest test type")
                        RunMultipleTest
                    Else
                        'MsgBox ("RunSingleTest test type")
                        RunSingleTest
                    End If
                End If
            End If
            If recirculation Then
                ' shut off everything associated with recirculation at end of
                ' test just in case they canceled in the middle of some flowing
                Move_Valve 12, "C"
                Move_Valve 19, "C"
                Move_Valve 20, "C"
                If v22_exists Then Move_Valve 21, "C"
                flow_status = 1
            End If
        Case 2
            RUNNING = False ' needs to be set this way so that manual control
                    ' realizes that it is not being called from within a test
            intest = False
            Pres% = 0
            HFLOW% = 0: lflow% = 0
            If lvperm_exclusive Then
                lv_man_ctrl.Show
                run_lv_man_ctrl
                ' 6.71.33
                'lv_man_ctrl.Hide
                Unload lv_man_ctrl
            Else
                If hasMultipleMVs Then
                    ManualControl1.Show 0
                Else
                    ManualControl.Show 0
                End If
                run_manual_control
            End If
            intest = True
        Case 3
            Screen.MousePointer = 11
            CleanOut
        Case 4
            LeakTest
        Case 5
            qcbubble.Show 1
        Case 6
            Aborted = True ' the default is true.  If they correctly set up a test and start it, this will be set to false
            cyclic_compression_setup.Show 1
            If Not Aborted Then run_cyclic_compression
        Case 7 ' vacuum purge
            do_purge_system
        Case 8 ' temperature log
            temperaturelog.Show 1
    End Select
    
click_done:
    mainmenu(1).Enabled = True
    mainmenu(5).Enabled = True
    Screen.MousePointer = 0
    Me.MousePointer = 0
    
End Sub

Private Sub selpen_Click(Index As Integer)

    penetrometer_select = Index
    selpen(Index).Checked = True
    selpen(3 - Index).Checked = False
    WPPS "Capstuff", "Penetrometer_Select", str$(penetrometer_select), CSFile$
    
End Sub

Private Sub selvolume_Click(Index As Integer)
    update_microflow_volume_selection_menu (Index)
End Sub

Public Sub update_microflow_volume_selection_menu(i As Integer)
    
    If i <> 4 Then
        Current_Microflow_Volume_Index = i
        selvolume(1).Checked = False
        selvolume(2).Checked = False
        selvolume(3).Checked = False
        selvolume(4).Checked = False
        selvolume(i).Checked = True
        microFlowUseAllVolumes = False
        WPPS "capstuff", "Current_Microflow_Volume_Index", Format$(i), CSFile$
        WPPS "capstuff", "microFlowUseAllVolumes", 0, CSFile$
    Else
        Current_Microflow_Volume_Index = 1
        selvolume(1).Checked = False
        selvolume(2).Checked = False
        selvolume(3).Checked = False
        selvolume(4).Checked = True
        microFlowUseAllVolumes = True
        WPPS "capstuff", "Current_Microflow_Volume_Index", Format$(1), CSFile$
        WPPS "capstuff", "microFlowUseAllVolumes", 1, CSFile$
    End If
End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False
    If simpleqc_enable Then
        setup_simple_qc_mode
    End If
    
End Sub

Public Sub LoadTextStrings()
' Load text elements for this form from external .ini file
    
    Dim i As Integer

    ' Form elements
    default_font.font = gpps2("default", "default_fontname", language$, "MS Sans Serif")
    default_font.fontsize = val(gpps2("default", "default_fontsize", language$, "8"))
    default_font.fontbold = (gpps2("default", "default_fontbold", language$, "Y") = "Y")
    system_font = default_font.font
    font_size = default_font.fontsize
    font_bold = default_font.fontbold
    
    GroupNmFrame.Caption = get_thing("main", "groupframe", language$, GroupNmFrame.Caption, GroupNmFrame, default_font)
    GroupNmFrame.fontsize = GroupNmFrame.fontsize + 2
    PressUnitFrame.Caption = get_thing("main", "pressunitframe", language$, PressUnitFrame.Caption, PressUnitFrame, default_font)
    PressUnitFrame.fontsize = PressUnitFrame.fontsize + 2
    lohmTableFrame.Caption = get_thing("main", "lohmtableframe", language$, lohmTableFrame.Caption, lohmTableFrame, default_font)
    lohmTableFrame.fontsize = lohmTableFrame.fontsize + 2
    GasFrame.Caption = get_thing("main", "gasframe", language$, GasFrame.Caption, GasFrame, default_font)
    GasFrame.fontsize = GasFrame.fontsize + 2
    set_fontstuff Label3, default_font
    Label3.fontsize = Label3.fontsize + 2
    set_fontstuff Label6, default_font
    Label6.fontsize = Label6.fontsize + 2
    set_fontstuff Label1, default_font
    Label1.fontsize = Label1.fontsize + 2
    set_fontstuff GasLabel, default_font
    GasLabel.fontsize = GasLabel.fontsize + 2
    mainmenu(0).Caption = gpps2("main", "mainmenu0", language$, mainmenu(0).Caption)
        For i = 1 To 4
            filemenu(i).Caption = gpps2("main", "filemenu" + str$(i), language$, filemenu(i).Caption)
        Next i
        exitmenu(0).Caption = gpps2("main", "filemenu15", language$, exitmenu(0).Caption)
        mode.Caption = gpps2("main", "mode", language$, mode.Caption)
        For i = 0 To 4
            modemenu(i).Caption = gpps2("main", "modemenu" + Format$(i), language$, modemenu(i).Caption)
        Next i
    mainmenu(1).Caption = gpps2("main", "mainmenu1", language$, mainmenu(1).Caption)
        For i = 1 To 4
            runmenu(i).Caption = gpps2("main", "runmenu" + str$(i), language$, runmenu(i).Caption)
        Next i
        runmenu(6).Caption = gpps2("main", "runmenu 6", language$, runmenu(6).Caption)
        runmenu(7).Caption = gpps2("main", "runmenu 7", language$, runmenu(7).Caption)
        chselect.Caption = gpps2("main", "chselect", language$, chselect.Caption)
        penselecthead.Caption = gpps2("main", "penselecthead", language$, penselecthead.Caption)
            selpen(1).Caption = gpps2("main", "selpen1", language$, selpen(1).Caption)
            selpen(2).Caption = gpps2("main", "selpen2", language$, selpen(2).Caption)
    mainmenu(2).Caption = gpps2("main", "mainmenu2", language$, mainmenu(2).Caption)
        reportmenu(1).Caption = gpps2("main", "report1", language$, reportmenu(1).Caption)
        reportmenu(2).Caption = gpps2("main", "report2", language$, reportmenu(2).Caption)
    mainmenu(3).Caption = gpps2("main", "mainmenu3", language$, mainmenu(3).Caption)
        For i = 0 To 4
            modifymenu(i).Caption = gpps2("main", "modify" + str$(i), language$, modifymenu(i).Caption)
        Next i
        psl(0).Caption = gpps2("main", "psl", language$, psl(0).Caption)
        prefs(0).Caption = gpps2("main", "prefs", language$, prefs(0).Caption)
    mainmenu(5).Caption = gpps2("main", "mainmenu5", language$, mainmenu(5).Caption)
        For i = 2 To 9
            calibmenu(i).Caption = gpps2("main", "calib" + str$(i), language$, calibmenu(i).Caption)
        Next i
        commdebugmenu.Caption = gpps2("main", "commdebug", language$, commdebugmenu.Caption)
        loadbackuplohm.Caption = gpps2("main", "loadbackup", language$, loadbackuplohm.Caption)
        'check_external_temperature.Caption = gpps2("main", "checkexttemp", language$, check_external_temperature.Caption)
    mainmenu(6).Caption = gpps2("main", "mainmenu6", language$, mainmenu(6).Caption)
        For i = 0 To 2
            helpmenu(i).Caption = gpps2("main", "help" + str$(i), language$, helpmenu(i).Caption)
        Next i
        
    ' Other text
    ts$(1) = gpps2("main", "ts1", language$, "Please open and empty the sample chamber(s).")
    ts$(2) = gpps2("main", "ts2", language$, "You don't have any chambers selected - Lohm calibration aborted")
    ts$(3) = gpps2("main", "ts3", language$, "and")
    ts$(4) = gpps2("main", "ts4", language$, "If the inlet pressure has changed")
    ts$(5) = gpps2("main", "ts5", language$, "the Flow Parameters Calibration must be run before")
    ts$(6) = gpps2("main", "ts6", language$, "the Lohm & Max Air Flow Calibration is run")
    ts$(7) = gpps2("main", "ts7", language$, "Calibrations")
    ts$(8) = gpps2("main", "ts8", language$, "Please insert the non-porous disk and seal the sample chamber.")
    ts$(9) = gpps2("main", "ts9", language$, "Setting Valve Limits")
    ts$(10) = gpps2("main", "ts10", language$, "Calibration of valve limits is now automatic.  You no longer need to manually run the valve limits calibration.")
    ts$(11) = gpps2("main", "ts11", language$, "Calibration of valve limits is complete.")
    ts$(12) = gpps2("main", "ts12", language$, "Calibrating flow meters")
    ts$(13) = gpps2("main", "ts13", language$, "Calibration of flow meters is complete.")
    ts$(14) = gpps2("main", "ts14", language$, "Calculating Lohm")
    ts$(15) = gpps2("main", "ts15", language$, "Abort Calibration")
    ts$(16) = gpps2("main", "ts16", language$, "Calibration of lohm value is complete.")
    ts$(17) = gpps2("main", "ts17", language$, "Calibration of max liquid flow complete.")
    ts$(18) = gpps2("main", "ts18", language$, "Calibration of diffusion volume complete.")
    ts$(19) = gpps2("main", "ts19", language$, "Calibration of regulator complete.")
    ts$(20) = gpps2("main", "ts20", language$, "Calibration of PC interface board complete.")
    ts$(21) = gpps2("main", "ts21", language$, "Calibration of pressure gauges is aborted.")
    ts$(22) = gpps2("main", "ts22", language$, "Calibration of pressure gauges complete.")
    ts$(23) = gpps2("main", "ts23", language$, "Calibration of compression regulator complete.")
    ts$(24) = gpps2("main", "ts24", language$, "Calibration Done")
'    ts$(25) = gpps2("main", "ts25", language$, "Enter Direct Comm Address")
    ts$(26) = gpps2("main", "ts26", language$, "You can not edit the DEFAULT group name.   Please switch to a different group or create a new group.")
    ts$(27) = gpps2("main", "ts27", language$, "Edit Group Name")
    ts$(28) = gpps2("main", "ts28", language$, "Please enter the new group name.")
    ts$(29) = gpps2("main", "ts29", language$, "New name already exists - name change aborted")
    ts$(30) = gpps2("main", "ts30", language$, "Something is wrong - current user index couldn't be found")
    ts$(31) = gpps2("main", "ts31", language$, "Enter Password")
    ts$(32) = gpps2("main", "ts32", language$, "Error changing name")
    ts$(33) = gpps2("main", "ts33", language$, "Error creating new user folder")
    ts$(34) = gpps2("main", "ts34", language$, "is already running in another window.")
    ts$(35) = gpps2("main", "ts35", language$, "Program Active")
    ts$(36) = gpps2("main", "ts36", language$, "Unable to locate")
    ts$(37) = gpps2("main", "ts37", language$, "file.  Please locate it.")
    ts$(38) = gpps2("main", "ts38", language$, "Please find the initialization file")
    ts$(39) = gpps2("main", "ts39", language$, "Unable to start CapWin without 'capwin.ini' file.  Check the Windows Explorer to locate the file.")
    ts$(40) = gpps2("main", "ts40", language$, "Exit CapWin and abort any running tests?")
    ts$(41) = gpps2("main", "ts41", language$, "Backup Lohm Calibration File")
    ts$(42) = gpps2("main", "ts42", language$, "There was an error copying the file.")
    ts$(43) = gpps2("main", "ts43", language$, "Enter New Password")
    ts$(44) = gpps2("main", "ts44", language$, "Error:  No chambers selected for testing")
    ts$(45) = gpps2("main", "ts45", language$, "Enter unit name: (6 char)")
    ts$(46) = gpps2("main", "ts46", language$, "The unit name must be less than 7 characters.")
    ts$(47) = gpps2("main", "ts47", language$, "New Unit")
    ts$(48) = gpps2("main", "ts48", language$, "Enter Conversion Factor")
    ts$(49) = gpps2("main", "ts49", language$, "Equivalent of 1 PSI in this new unit.")
    ts$(50) = gpps2("main", "ts50", language$, "i.e. 1 PSI = 0.0689 BAR")
    ts$(51) = gpps2("main", "ts51", language$, "The unit factor must be positive")
   ' ts$(52) = gpps2("main", "ts52", language$, "PSI")
   ' ts$(53) = gpps2("main", "ts53", language$, "KPA")
   ' ts$(54) = gpps2("main", "ts54", language$, "kg/cm2")
   ' ts$(55) = gpps2("main", "ts55", language$, "BAR")
   ' ts$(56) = gpps2("main", "ts56", language$, "in")
   ' ts$(57) = gpps2("main", "ts57", language$, "cm")
   ' ts$(58) = gpps2("main", "ts58", language$, "mm")
   ts$(59) = gpps2("main", "ts59", language$, "Error:  Your hardware is not compatible with this software - performance counter hardware not detected")
   ts$(60) = gpps2("main", "ts60", language$, "Error: No communication with the external temperature controller")
   ts$(61) = gpps2("main", "ts61", language$, "External temperature")
   ts$(62) = gpps2("main", "ts62", language$, "Piston about to lower")
   ts$(63) = gpps2("main", "ts63", language$, "Piston about to rise")
   ts$(64) = gpps2("main", "ts64", language$, "End Calibration")
   ts$(65) = gpps2("main", "ts65", language$, "Pick a place to save this new lohm table.")
   ts$(66) = gpps2("capflow", "ts439", language$, "Save New Lohm Table")
End Sub

Private Sub viewReportMenu_Click(Index As Integer)
    Debug.Print "Open report option"
    CommonDialog1.DialogTitle = "Open .CFT file"
    CommonDialog1.Filter = "CAPWIN Data file (*.cft)"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
                
    On Error Resume Next
    CommonDialog1.ShowOpen
            
    If Err Then
        MsgBox "Dialog Cancelled"
        Exit Sub
    End If
    Dim rval As Single
    rval = Shell(EXE_Path$ + "Caprep.exe " + "resultcall " + CommonDialog1.filename, 1)
End Sub
Private Sub testGas()
    'Purpose to open v1 and v2, then turn on reg1 and make sure flow/pressure increase.
    'Then close v1 and v2 and zero reg and wait until values settle or a define time peroid.
    
    'Dim startPressure As Single
    'Dim endPressure As Single
    'Dim currentPressure As Single
    'Dim currentRange As Single
    Dim startHighFlow As Single
    Dim startLowFlow As Single
    Dim endHighFlow As Single
    Dim currentHighFlow As Single
    Dim endLowFlow As Single
    Dim currentLowFlow As Single
    Dim currentTime As Single
    Dim startTime As Single
    Dim incTime As Single
    
    Dim gasConnected As Boolean
    Dim userQuit As Boolean
    
    Dim currentLowFlowRange As Integer
    Dim currentHighFlowRange As Integer
    
    currentLowFlowRange = lflow%
    lflow% = 0
    
    currentHighFlowRange = HFLOW%
    If xhflow Then
        HFLOW% = 2
    Else
        HFLOW% = 1
    End If
    
    'read start pressure
    ReadXReturnX4 0
    startLowFlow = x5
    
    'read start flow
    ReadXReturnX4 1
    startHighFlow = x5
    
    If Not nov2 Then
        'open v2
        SetV2Limits 'just in case
        OPENV2MOSTLY
    End If
    
    'open v3
    Move_Valve 2, "O"
    
    If low_flow_controller Then
        'open v1
        Move_Valve 0, "O"
        zero_lfctrl
        inc_lfctrl 2000
    End If
    'read pressure and flow until they increase by 10%
    gasConnected = False
    status.Label1.Caption = "Testing gas connection"
    status.Label2.Caption = "Reading Flow Rates - Please Wait"
    status.Label3.Caption = ""
    status.Show
    
   
    Zero_Reg
    inc_reg 100
    startTime = Timer
    incTime = 1
    While Not gasConnected And status.Command1.Enabled = True
        currentTime = Timer - startTime
        'read low flow
        ReadXReturnX4 0
        currentLowFlow = x5
              
        'read high flow
        ReadXReturnX4 1
        currentHighFlow = x5
        
        'update screen
        status.Label2.Caption = "Change in Low Flow:" + Xformat(currentLowFlow, "###0.0##  ") + "SCCM"
        status.Label3.Caption = "Change in High Flow:" + Xformat(currentHighFlow, "###0.0##  ") + "SCCM"
        status.Label2.Refresh
        status.Label3.Refresh
                
        If currentTime > incTime Then
            'increment regulator
            inc_reg 100
            incTime = incTime + currentTime
        End If
        'determine if gas is connected
        If currentHighFlow > 1000 Or currentLowFlow > 1 Then
            gasConnected = True
        ElseIf REGPOS > 3000 Then
            If MsgBox("Ensure that the gas is connected and turned on. Press Ok to continue or Cancel to quit", vbOKCancel) = vbCancel Then
                userQuit = True
            End If
        End If
        
        
    Wend
    'unload status window
    status.Hide
    
    'reset abort button if needed
    If status.Command1.Enabled = False Then
        status.Command1.Enabled = True
    End If
    
    'zero regulator
    Zero_Reg
    
    'zero flow controller
    zero_lfctrl
    
    'reposition v2
    close_v2_completely
    
    'close v1
    Move_Valve 0, "C"
        
    waitseconds 5
    
    Move_Valve 2, "C"
    
    reset_status_window
        
    lflow% = currentLowFlowRange
    
    HFLOW% = currentHighFlowRange
       
    If userQuit Then
        Exit Sub
    End If
    
    
End Sub

