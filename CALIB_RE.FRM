VERSION 5.00
Begin VB.Form calib_reg 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regulator Calibration"
   ClientHeight    =   3480
   ClientLeft      =   1980
   ClientTop       =   2490
   ClientWidth     =   5205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "CALIB_RE.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3480
   ScaleWidth      =   5205
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "&Start"
         Height          =   492
         Left            =   240
         TabIndex        =   1
         Top             =   2640
         Width           =   2052
      End
      Begin VB.Frame MsgFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "NOTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   4332
         Begin VB.Label lblPaused 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "PAUSED"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   480
            TabIndex        =   6
            Top             =   1800
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Make sure that the pressurized air hose is attached and the instrument is turned on."
            ForeColor       =   &H80000008&
            Height          =   1320
            Left            =   600
            TabIndex        =   5
            Top             =   480
            Width           =   3180
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   600
            TabIndex        =   4
            Top             =   1320
            Width           =   3132
         End
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   492
         Left            =   2520
         TabIndex        =   2
         Top             =   2640
         Width           =   2052
      End
   End
   Begin VB.Menu mnuPause 
      Caption         =   "Pause"
   End
   Begin VB.Menu mnuManualControl 
      Caption         =   "Manual Control"
   End
End
Attribute VB_Name = "calib_reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ts$(9)                      ' Text strings for this form
Dim runCount As Integer         ' Keep track of runs for dual-regulator system
Dim pauseRegCalib As Boolean

Private Sub Command1_Click()

    Dim reg_been_over As Boolean, last_t!, T!, i As Long, t1!, t2!, t3!, n%
    Dim p_atm As Single
    Dim fn1 As Integer, fn2 As Integer
    Dim rcl As Long, rol As Long
    Dim regposchan As Integer
    Dim presschan As Integer
    Dim temp_maxpres
    Dim current_percent As Single
    Dim percent_at_zero As Single

    Command1.Enabled = False
    pauseRegCalib = False
    mnuPause.Enabled = True
    mnuManualControl.Enabled = False
    lblPaused.Visible = False
    'Maxpres = 500 : debugging - AW 1/11/2016
    percent_at_zero = 0
    If compregcal Then
        If ip_creg_enable Then
            rcl = 0
            rol = 4000
            regposchan = -2 ' there is no compressure regulator position channel, remap to CREGPOS
        Else
            rcl = raw_reading(32)
            rol = raw_reading(33)
            regposchan = 31
        End If
        presschan = 35
    ElseIf ip_reg_enable Then
        rcl = 0
        rol = 4000
        presschan = 2
        regposchan = -1 ' there is no regulator position channel, remap to REGPOS
    Else
        rcl = reg_cl
        rol = reg_ol
        regposchan = 18
        presschan = 2
    End If
    
    If dualregulator Then
        temp_maxpres = Maxpres
        If Maxpres > reg1pmax Then
            Maxpres = reg1pmax
        End If
        Move_Valve 16, "C"
        regnum = 0
        If (air_inlets = 2) Then
            Move_Valve 15, "C"
            current_air_inlet = 1
        End If

    End If
restart_regcal:

    While pauseRegCalib
        DoEvents
    Wend
    
    If air_inlets > 1 And (compregcal Or (dualregulator And Vpos(17) = 1)) Then
        Move_Valve 15, "O"
        current_air_inlet = 2
    End If
    reg_been_over = False
    n% = 0
    MsgFrame.Caption = ""
    Label1.Caption = ts$(1)         ' "Initializing Hardware"
    Label2.Caption = ""
    
    If compregcal Then
        zero_creg
        p_atm = 14.7
    Else
        Zero_Reg
        Move_Valve 10, "C"
        
        'AJB 11-01-09
        If cartridge_tester Then
            Move_Valve 5, "C"
        End If
        'AJB 11-01-09
        
        Pres% = 0
        ReadXReturnX4 2
        p_atm = x5
        ' seal the system now, after getting the atmospheric pressure
        If H2OPERM Then
            Dry_Chamber_Control "C"
        End If
        Move_Valve 2, "C"
        Rem open valve 2 all the way
        OPENV2MOSTLY
        RUNNING = True ' allows auto-calibration of crossover
    End If
    
    last_t! = 0
    fn1 = FreeFile
    Open EXE_Path$ + "capwin.tmp" For Output As #fn1
    Rem create file of regulator position (18) and pressure (in PSIA)
    Rem for compregcal, position is 31 and pressure is 35
    Rem for I/P converter, regulator position is variable REGPOS
    MsgFrame.Caption = ""

    Do
        While pauseRegCalib
            DoEvents
        Wend
        T! = Timer
        
        Do
            While pauseRegCalib
                DoEvents
            Wend
        
            If regposchan = -1 Then
                i = REGPOS
            ElseIf regposchan = -2 Then
                i = CREGPOS
            Else
                i = raw_reading(regposchan)
            End If
            
            If i <= rcl Then
                x5 = 0
            ElseIf i >= rol Then
                x5 = 100
                reg_been_over = True
            Else
                x5 = 100# * (i - rcl) / (rol - rcl)
            End If
            
            Label1.Caption = ts$(2) + ": " + Format$(x5, "##0.#")    ' "Regulator Position"
            current_percent = x5
            ReadXReturnX4 presschan
            t1! = x5
            Label2.Caption = ts$(3) + ": " + Xformat$((t1! - p_atm) * PCNV, "###0.000  ") + PU$     ' "Pressure"
            DoEvents
            ReadXReturnX4 presschan
            t2! = x5
            Label2.Caption = ts$(3) + ": " + Xformat$((t2! - p_atm) * PCNV, "###0.000  ") + PU$     ' "Pressure"
            DoEvents
            ReadXReturnX4 presschan
            t3! = x5
            Label2.Caption = ts$(3) + ": " + Xformat$((t3! - p_atm) * PCNV, "###0.000  ") + PU$     ' "Pressure"
            DoEvents
            If Command2.Enabled = False Then Exit Do
            If (Timer - T!) > 4 Then
                t1! = (t1! + t2! + t3!) / 3
                Exit Do
            End If
        Loop Until t1! = t2! And t2! = t3!
        
        If t1! - p_atm > last_t! Then
            Print #fn1, i; ","; str$(t1! - p_atm)
            n% = n% + 1
            last_t! = t1! - p_atm
        End If
        
        If compregcal Then
            If t1! >= PY2(6) Then
                reg_been_over = False
                Exit Do
            End If
        Else
            If t1! >= PY2(0) Then
                reg_been_over = False
                Exit Do
            End If
            If t1! - p_atm > Maxpres Then
                reg_been_over = False
                Exit Do
            End If
        End If
        
        If Command2.Enabled = False Then Exit Do
        If reg_been_over Then Exit Do
        ' at less than 2 PSI, we increment in really small amounts
        If (t1! - p_atm) < 2 Then
            percent_at_zero = current_percent
            GoSub inc_really_small
        ElseIf current_percent - percent_at_zero < 5 Then
            manrunning = True
            GoSub inc_really_small
            manrunning = False
            GoSub inc_really_small
        Else
            ' set manrunning flag so first incs don't wait
            ' for regulator to stop
            manrunning = True
            ' at less than 5%, we use smallest increment
            ' increase by more increments up to above 25% when we use full speed
            If current_percent - percent_at_zero > 10 Then
                GoSub inc_thing
                If current_percent > 15 Then
                    GoSub inc_thing
                    If current_percent > 20 Then
                        GoSub inc_thing
                        If current_percent > 25 Then
                            GoSub inc_thing
                        End If
                    End If
                End If
            End If
            manrunning = False
            GoSub inc_thing
        End If
       
    Loop

    Close #fn1
    MsgFrame.Caption = ""
    Label1.Caption = ts$(4)     ' "Zeroing system"
    Label2.Caption = ""
    
    If air_inlets > 1 And current_air_inlet = 2 Then
        Move_Valve 15, "C"
        current_air_inlet = 1
    End If
    
    If compregcal Then
        zero_creg
    Else
        Zero_Reg
        excercize_valve_2
        Move_Valve 2, "O"
        If auto_wet_enable = False Then
            Move_Valve 11, "O"
        End If
    End If
    
    If Command2.Enabled = False Then
        MsgBox ts$(5)           ' "Calibration Stopped by user - you may want to recalibrate"
    ElseIf reg_been_over And (dualregulator = False Or Vpos(17) = 1) Then
        MsgBox ts$(6) + vbCrLf + ts$(7) + ":" + Xformat$((last_t!) * PCNV, "###0.000  ") + PU$     ' "Note: Check air inlet pressure - calibration did not reach maximum pressure"/"Maximum pressure seen"
    End If
    
    fn1 = FreeFile
    Open EXE_Path$ + "capwin.tmp" For Input As #fn1
    fn2 = FreeFile
    If compregcal Then
        Open EXE_Path$ + "capwincompreg.cal" For Output As #fn2
    ElseIf dualregulator And Vpos(17) = 1 Then
        regnum = 1
        Open EXE_Path$ + "capwinrg2.cal" For Output As #fn2
    Else
        regnum = 0
        Open EXE_Path$ + "capwinrg.cal" For Output As #fn2
    End If
    Print #fn2, n%
    If compregcal Then
        creg_table_size% = n%
        ReDim creg_table_pos(n%) As Long, creg_table_pres!(n%)
        For i = 1 To n%
            DoEvents
            Input #fn1, creg_table_pos(i), creg_table_pres!(i)
            Print #fn2, creg_table_pos(i); ","; str$(creg_table_pres!(i))
        Next i
    ElseIf regnum = 0 Then
        reg_table_size%(0) = n%
        ReDim reg_table_pos(n%) As Long, reg_table_pres!(n%)
        For i = 1 To n%
            DoEvents
            Input #fn1, reg_table_pos(i), reg_table_pres!(i)
            Print #fn2, reg_table_pos(i); ","; str$(reg_table_pres!(i))
        Next i
    Else
        reg_table_size%(1) = n%
        ReDim reg_table_pos2(n%) As Long, reg_table_pres2!(n%)
        For i = 1 To n%
            DoEvents
            Input #fn1, reg_table_pos2(i), reg_table_pres2!(i)
            Print #fn2, reg_table_pos2(i); ","; str$(reg_table_pres2!(i))
        Next i
    End If
    Close #fn1
    Close #fn2
    Kill EXE_Path$ + "capwin.tmp"
    If regnum = 0 And dualregulator And compregcal = False Then
        Move_Valve 2, "C" ' close the vents as we are going up again
        Move_Valve 11, "C"
        Move_Valve 16, "O"
        regnum = 1
        If (air_inlets = 2) Then
            Move_Valve 15, "O"
        '   waitseconds 2
            current_air_inlet = 2
        End If

        Maxpres = temp_maxpres
        GoTo restart_regcal
    End If
    Command2.Enabled = True
    Command2.Caption = "Done"
    While Command2.Enabled
        DoEvents
    Wend
    
    calibrationComplete = True
    Unload Me
'Else
'    command1.enabled = False
'End If
Exit Sub

inc_really_small:
    If compregcal Then
        inc_creg 5
    Else
        inc_reg 5
    End If
    Return

inc_thing:
    If compregcal Then
        inc_creg 20
    Else
        inc_reg 20
    End If
    Return

End Sub

Private Sub Command2_Click()

If Command1.Enabled Then
    Unload Me
Else
    Command2.Enabled = False
End If

End Sub

Private Sub Form_Load()

    LoadTextStrings
    If compregcal Then
        Me.Caption = ts$(8)         ' "Compression Regulator Calibration"
    End If
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(9) + str$(unitnumber)    ' "Unit"
    End If
'    Command1.SetFocus
    runCount = 0
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    calib_reg.Caption = gpps2("calib_reg", "window title", language$, calib_reg.Caption)
    MsgFrame.Caption = get_thing("calib_reg", "msgframe", language$, MsgFrame.Caption, MsgFrame, default_font)
    Label1.Caption = get_thing("calib_reg", "label1", language$, Label1.Caption, Label1, default_font)
    Command1.Caption = gpps2("calib_reg", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    Command2.Caption = gpps2("calib_reg", "command2", language$, Command2.Caption)
    set_fontname Command2, default_font
    Label2.Caption = get_thing("calib_reg", "label2", language$, Label2.Caption, Label2, default_font)

    ' Other text
    ts$(1) = gpps2("calib_reg", "ts1", language$, "Initializing Hardware")
    ts$(2) = gpps2("calib_reg", "ts2", language$, "Regulator Position")
    ts$(3) = gpps2("calib_reg", "ts3", language$, "Pressure")
    ts$(4) = gpps2("calib_reg", "ts4", language$, "Zeroing system")
    ts$(5) = gpps2("calib_reg", "ts5", language$, "Calibration stopped by user - you may want to recalibrate")
    ts$(6) = gpps2("calib_reg", "ts6", language$, "Note: Check air inlet pressure - calibration did not reach maximum pressure")
    ts$(7) = gpps2("calib_reg", "ts7", language$, "Maximum pressure seen")
    ts$(8) = gpps2("calib_reg", "ts8", language$, "Compression Regulator Calibration")
    ts$(9) = gpps2("calib_reg", "ts9", language$, "Unit")
    
End Sub

Private Sub mnuManualControl_Click()
    RUNNING = True
    If hasMultipleMVs Then
        ManualControl1.Show 0
    Else
        ManualControl.Show 0
    End If
    run_manual_control
End Sub

Private Sub mnuPause_Click()
    If mnuPause.Caption = "Pause" Then
        lblPaused.Visible = True
        mnuPause.Caption = "Resume"
        mnuManualControl.Enabled = True
        pauseRegCalib = True
    Else
        lblPaused.Visible = False
        mnuPause.Caption = "Pause"
        mnuManualControl.Enabled = False
        pauseRegCalib = False
    End If
End Sub
