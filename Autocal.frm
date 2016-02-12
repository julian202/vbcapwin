VERSION 5.00
Begin VB.Form Autocal 
   BackColor       =   &H000000FF&
   Caption         =   "Automatic Gauge Calibration"
   ClientHeight    =   5940
   ClientLeft      =   1440
   ClientTop       =   1890
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9630
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.TextBox Text2 
         Height          =   4575
         Left            =   4560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "Autocal.frx":0000
         Top             =   240
         Width           =   4695
      End
      Begin VB.ComboBox gaugelist 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3840
         Top             =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Read All Values"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4200
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Run PG Cal"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Run FM Cal"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   2175
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   10
         Left            =   5280
         Max             =   255
         Min             =   1
         TabIndex        =   14
         Top             =   5280
         Value           =   1
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Stop"
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Adjust Zero"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Adjust Span"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   5160
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Don't change ini file"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Calibrating low pressure gauge: 5 PSI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label5 
         Caption         =   "Current reading:"
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
         TabIndex        =   16
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Cal:"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Uncal:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lowreading 
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
         Left            =   1560
         TabIndex        =   4
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label highreading 
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
         Left            =   1560
         TabIndex        =   6
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   5280
         TabIndex        =   15
         Top             =   4920
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Autocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Test module for automatic calibration of pressure gauges

' 1) Vent to atmospheric pressure
' 2) Read gauge and adjust zero point
' 3) Set up values for pressurization
' 4) Get target pressure; go to target
' 5) Click-record gauge; input calibrated correspondence
' 6) Repeat 3-4 until done
' 7) Calculate new .ini file values and adjust params in ini file
' 8) Repeat calibration to double-check

Option Explicit

Dim flowstring$, pressurestring$
Dim readingIndex$(10)                   ' List of transducers that can be read
Dim targetGauge$                        ' Current transducer being read
Dim results() As Single                 ' Results of the readings. First column is for
                                        ' subject values, second is for calibrated values
Dim settingList() As Single             ' List of pressure or flow settings at which readings are taken
                                        ' Beyond first value of zero
Dim statstring$                         ' text2 message string
Dim abort As Boolean                    ' Flag to stop the current procedure
Dim numreadings As Integer              ' Number of readings to take
Dim cal_done As Boolean                 ' Flag that the calibration is complete
Dim ts$(32)                             ' Text strings for this form


Private Sub Command1_Click()
' Run a pressure gauge calibration
    
    Dim maxrange As Single
    Dim tempreturn$
    Dim temprange As Single, increment As Single
    Dim channel As Integer
    Dim u_low As Integer, u_high As Integer   ' pres% values for uncal low & high ranges
    Dim c_low As Integer, c_high As Integer     ' pres% values for cal low & high ranges
    Dim c_avg As Single, u_avg As Single        ' Averaging values
    Dim at_target As Boolean
    Dim i, j, k As Integer
    Dim minrange As Single                  ' Adjustment for differential gauge
   ' Dim valvepercent As Single              ' setting for valve 2
    Dim real_vals(3) As Single              ' temp storage for vars of flow meter
    Dim span_multiplier As Single, span As Single
    Dim crossover_reading As Single
    Dim meter_index As Integer, alt_index As Integer    ' indices for accessing parameter values
'Dim debugstring$
    Dim temp$
    Dim lowrange As Boolean             ' flag to see if we're in low or high range
    Dim temp_diff As Single

    cal_done = False
    abort = False
' singlepointgastestShow 0
' Procedure:

' 0) Set up ranges, v2 positions, and other settings depending on the flow meter chosen
    If targetGauge$ = "HPG" Then
        ' Range select
        maxrange = PY2(0)
        minrange = PY1(0)
      '  valvepercent = 10
        meter_index = 0: alt_index = 2
    ElseIf targetGauge$ = "LPG" Then
        ' Range select
        maxrange = PY2(2) - 0.5     ' Don't want to get too close to upper limit of LPG
        minrange = PY1(2)
  '      valvepercent = 5
        meter_index = 2: alt_index = 0
        Check1.value = vbUnchecked          ' Don't need to change max pressure for LPG
    End If
    
    ' Store the limit values of the gauge not being used
    real_vals(0) = PY1(alt_index)
    real_vals(1) = PY2(alt_index)
    real_vals(2) = PY1(alt_index + 1)
    real_vals(3) = PY2(alt_index + 1)
    
    ' Make the unused gauge an "ideal" pressure gauge of the correct range.
    PY1(alt_index) = InputBox(ts$(1), , "0")        ' "Enter min. value of calibrated gauge in PSI. (e.g. 14.7 for differential, 0 for absolute)"
    PY1(alt_index + 1) = PY1(alt_index)
    PY2(alt_index) = InputBox(ts$(2), , "110")      ' "Enter max. value of calibrated gauge in PSI."
    PY2(alt_index + 1) = PY1(alt_index) + 0.2 * (PY2(alt_index) - PY1(alt_index))
   ' If targetGauge$ = "HPG" Then
   '     PY1(alt_index) = 0
  '      PY2(alt_index) = 110
   '     PY1(alt_index + 1) = 0
   '     PY2(alt_index + 1) = 22
   ' ElseIf targetGauge$ = "LPG" Then
   '     PY1(alt_index) = 14.7
   '     PY2(alt_index) = 19.7
   '     PY1(alt_index + 1) = 14.7
   '     PY2(alt_index + 1) = 15.7
   ' End If
    
' 1) Figure out points to take (besides zero)

    While val(tempreturn$) < 4
        tempreturn$ = InputBox(ts$(3), , 4)         ' "Enter the number of readings to take (min. 4)"
        If tempreturn$ = "" Then Exit Sub
    Wend
    numreadings = val(tempreturn$)
    ReDim settingList(numreadings)
    ReDim results(numreadings, 2)
  '  If Check1.Value = vbChecked Then        ' Prompt for new maximum pressure
        tempreturn$ = InputBox(ts$(4), , str$(maxrange))        ' "Enter a maximum pressure for the calibration"
        maxrange = val(tempreturn$)
    'End If
    settingList(1) = (maxrange - minrange) / 5 + minrange
    settingList(2) = -1                 ' indicates crossover from low range to high
    If numreadings = 4 Then
        settingList(3) = maxrange
    Else
        temprange = maxrange - settingList(1)   ' gives the distance from the crossover point to max pressure
        increment = temprange / (numreadings - 3)
        For i = 3 To numreadings - 1
            If settingList(i - 1) = -1 Then
                settingList(i) = settingList(i - 2) + increment
            Else
                settingList(i) = settingList(i - 1) + increment
            End If
        Next i
    End If
    
' 2) Set valves and zero regulator
    Dry_Chamber_Control "O"
    If ExtraPG And Not targetGauge$ = "LPG" Then Move_Valve 10, "C"     ' Protect LPG
    If targetGauge$ = "LPG" Then Move_Valve 10, "O"
    Zero_Reg
    Move_Valve 13, "O"
    Move_Valve 13, "C"
    
' 2a) Set values for readings
    channel = 2                         ' Pressure gauge
    If targetGauge$ = "HPG" Then
        u_low = 1: u_high = 0
        c_low = 3: c_high = 2
    Else
        u_low = 3: u_high = 2
        c_low = 1: c_high = 0
    End If
    
' 3) Take zero point
    Text2.Text = ts$(5)     ' "At zero point, waiting for stability"
    V2POS = cLimit
    Move2V2Pos
    waitseconds 10
    If abort Then cal_aborted: Exit Sub
    DoEvents
'    tempreturn$ = ""
'    While tempreturn$ = ""
'        tempreturn$ = InputBox("Enter calibrated pressure value:")
'    Wend

    u_avg = 0: c_avg = 0
    For k = 1 To 10
    
        ' Read uncalibrated
        Pres% = u_low                        ' Set to low range
        ReadXReturnX4 channel
        u_avg = u_avg + x5
       ' results(1, 2) = Val(tempreturn$)
    
        
        ' Read calibrated
        Pres% = c_low
        ReadXReturnX4 channel
        c_avg = c_avg + x5
    Next k

    results(1, 1) = u_avg / 10
    results(1, 2) = c_avg / 10
    
    ' Correct the zero point. This is a little different than for the flow meters:
    ' If an absolute gauge (i.e. min value is 0) we want to adjust zero so that reading
    ' matches the calibrated gauge. If absolute, we do want the initial reading to be 0.
    If PY1(meter_index) < 10 Then        ' absolute, assuming not a vacuum system
        temp_diff = results(1, 2) - results(1, 1)
    Else
        temp_diff = 0 - results(1, 1)
    End If
    
    PY1(meter_index) = PY1(meter_index) + temp_diff
    PY2(meter_index) = PY2(meter_index) + temp_diff
    PY1(meter_index + 1) = PY1(meter_index + 1) + temp_diff
    PY2(meter_index + 1) = PY2(meter_index + 1) + temp_diff
        
    
    statstring$ = ts$(6) + ":" + vbCrLf + Format$(results(1, 1), "##0.00") + " PSI " + ts$(7) + vbCrLf + Format$(results(1, 2), "##0.00") + " PSI " + ts$(8) + vbCrLf ' Point taken/uncalibrated/calibrated
    statstring$ = statstring$ + ts$(9) + " =" + Format$(Abs((results(1, 2) - results(1, 1)) / results(1, 2)) * 100, "###0.00") + "%" + vbCrLf       ' "Difference"
    Text2.Text = statstring$

    lowrange = True
' For each additional point:
    For i = 1 To numreadings - 1
    
        If abort Then cal_aborted: Exit Sub
        
        If lowrange Then Pres% = c_low Else Pres% = c_high
        
        ReadXReturnX4 channel
        If x5 >= settingList(i) Then at_target = True Else at_target = False
        V2POS = cLimit + (oLimit - cLimit) * V2Percent / 100                       ' Open to 10%
        OpenV2Pos

       ' Pres% = c_low       ' start off in low range until we hit crossover
        While Not at_target

        '    debugstring = ""
        
            If abort Then cal_aborted: Exit Sub
            
            ' 4) Open V2 and regulator until pressure is achieved
            Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(10) + " ..."   ' "Reaching next pressure"
            If targetGauge$ = "HPG" And Not newreg Then
                For j = 1 To 10
                    inc_reg HScroll1.value
                Next j
            Else
                inc_reg HScroll1.value
            End If
            
            ' 5) Wait for stability
            Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(11) + " ..."       ' "Waiting for stability"
            waitseconds 5
            
            Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(12) + " ..."       ' "Reading value"
            If lowrange And settingList(i) = "-1" Then
                lowrange = False      ' At crossover point
                Pres% = c_high
                settingList(i) = settingList(i - 1)
            End If
            
            ReadXReturnX4 channel

            If x5 >= settingList(i) Then at_target = True
    '        debugstring = debugstring + "u=" + Str$(X5) + vbCrLf
            ReadXReturnX4 channel
    '        debugstring = debugstring + "c=" + Str$(X5) + vbCrLf
    '        debugstring = debugstring + "i=" + Str$(i) + vbCrLf
    '        debugstring = debugstring + Str$(Now) + vbCrLf
    '        debugstring = debugstring + Str(Len(debugstring))
    '        singlepointgastestText1.Text = debugstring
        Wend
        
            If lowrange And settingList(i) = "-1" Then
                lowrange = False      ' At crossover point
                Pres% = c_high
                settingList(i) = settingList(i - 1)
            End If
            
        Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(13)        ' "Pressure reached"
        V2POS = cLimit
        Move2V2Pos
        waitseconds 5
        
        ' 6) Prompt user for calibrated value
     '   tempreturn$ = ""
      '  While tempreturn$ = ""
     '       tempreturn$ = InputBox("Enter calibrated pressure value:")
     '   Wend
        
        ' 7) Store along with recorded value
        ' Uncalibrated
        u_avg = 0: c_avg = 0
        For k = 1 To 10
            If lowrange Then Pres% = u_low Else Pres% = u_high
            ReadXReturnX4 channel
            u_avg = u_avg + x5
     '   results(i, 2) = Val(tempreturn$)
        
            'Calibrated
            If lowrange Then Pres% = c_low Else Pres% = c_high
            ReadXReturnX4 channel
            c_avg = c_avg + x5
        Next k
        
        results(i, 1) = u_avg / 10
        results(i, 2) = c_avg / 10
        
        
        statstring$ = statstring$ + vbCrLf + ts$(6) + ":" + vbCrLf + Format$(results(i, 1), "##0.00") + " PSI " + ts$(7) + vbCrLf + Format$(results(i, 2), "##0.00") + " PSI " + ts$(8) + vbCrLf    ' "Point taken"/"uncalibrated"/"calibrated"
        statstring$ = statstring$ + ts$(9) + " =" + Format$(Abs((results(i, 2) - results(i, 1)) / results(i, 2)) * 100, "###0.00") + "%" + vbCrLf       ' "Difference"
    Next i
    
    Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(14)        ' "Pressure readings complete"
    
' 8) Eventually, do calculations and change values at end
    V2POS = cLimit + (oLimit - cLimit) * V2Percent / 100
    OpenV2Pos
    Zero_Reg
    Move_Valve 13, "O"
    V2POS = cLimit
    Move2V2Pos
    
    do_calcs
    
        ' Adjust the low-range offset so that the high and low ranges match at the
    ' crossover point. The crossover point value for the high range is stored in
    ' results(1,1); for the low range in crossover_reading
 '   FY1(1, meter_index + 1) = FY1(1, meter_index + 1) + (results(1, 1) - crossover_reading)
 '   fy2(1, meter_index + 1) = fy2(1, meter_index + 1) + (results(1, 1) - crossover_reading)
    
    ' Now work out the span correction for both high and low ranges
    span_multiplier = results(numreadings - 1, 2) - results(numreadings - 1, 1) ' cal-uncal
    span_multiplier = span_multiplier / results(numreadings - 1, 2)
  '  MsgBox (Str$(span_correct))
    ' if span_correct is positive, span has to be increased. Otherwise, decreased
    
    span_multiplier = 1 + span_multiplier
    
    ' Span is the difference between the zero point and max point
    span = PY2(meter_index) - PY1(meter_index)
    
    If Check1.value = vbUnchecked Then
        MsgBox (ts$(15) + " " + str$(span) + " " + ts$(16) + " " + str$(span * span_multiplier)) ' "Span will be changed from"/"to"
    End If
    
    ' Reset the values for the "other" meter
    PY1(alt_index) = real_vals(0)
    PY2(alt_index) = real_vals(1)
    PY1(alt_index + 1) = real_vals(2)
    PY2(alt_index + 1) = real_vals(3)
    
    If Check1.value = vbUnchecked Then
        ' Set new span values here
        PY2(meter_index) = (span * span_multiplier) + PY1(meter_index)
        PY2(meter_index + 1) = (span * span_multiplier * 0.2) + PY1(meter_index + 1)
        
        ' Write the new values out
        WPPS "Capstuff", "PY1_" + Format(meter_index), Format(PY1(meter_index)), CSFile$
        WPPS "Capstuff", "PY2_" + Format(meter_index), Format(PY2(meter_index)), CSFile$
        WPPS "Capstuff", "PY1_" + Format(meter_index + 1), Format(PY1(meter_index + 1)), CSFile$
        WPPS "Capstuff", "PY2_" + Format(meter_index + 1), Format(PY2(meter_index + 1)), CSFile$
    End If

End Sub

Private Sub Command2_Click()
' Run a flow meter calibration
    
    Dim maxrange As Single
    Dim tempreturn$
    Dim temprange As Single, increment As Single
    Dim channel As Integer
    Dim c_low As Integer, c_high As Integer   ' pres% values for low & high ranges of cal. gauge
    Dim u_low As Integer, u_high As Integer   ' pres% values for low & high ranges of uncal. gauge
    Dim v_cal As Integer, v_uncal As Integer   ' settings for v_flow
    Dim u_avg, c_avg As Single                  ' Reading averaging values
    Dim meter_index As Integer, alt_index As Integer    ' indices for FY1(1,x), etc.
    Dim at_target As Boolean
    Dim i, k As Integer
    Dim valvepercent As Single              ' setting for valve 2
    Dim value_multiplier As Integer         ' Multiply inputs by different values depending on range
    Dim zero_cal As Single, zero_uncal As Single
    Dim real_vals(3) As Single              ' temp storage for FY vars of flow meter not being used
    Dim span_multiplier As Single, span As Single
    Dim crossover_reading As Single
    Dim lowhigh_select As Integer           ' select lfm or high/xfm
    
    cal_done = False
    abort = False
    
' Procedure:
' 0) Set up ranges, v2 positions, and other settings depending on the flow meter chosen
    lowhigh_select = 1      ' for high/xfm
    If targetGauge$ = "XFM" Then
        ' Range and V2 select
        maxrange = FY2(1, 2)
        valvepercent = 80
        value_multiplier = 1
        ' Calibrating the highest range implies that the lower range is now connected to a
        ' calibrated gauge. Set the ini values of the HFM to those of a "perfect" XFM and
        ' save the actual values for later.
        meter_index = 2: alt_index = 0      ' HFM: FY1(1,0), XFM:FY1(1,2)
    ElseIf targetGauge$ = "HFM" Then
        ' Range and V2 select
        maxrange = FY2(1, 0)
        valvepercent = V2Percent
        value_multiplier = 1000
        ' Calibrating the middle range implies that the higher range is now connected to a
        ' calibrated gauge. Set the ini values of the XFM to those of a "perfect" HFM and
        ' save the actual values for later.
        meter_index = 0: alt_index = 2      ' HFM: FY1(1,0), XFM:FY1(1,2)
    ElseIf targetGauge$ = "LFM" Then
        maxrange = FY2(0, 0)
        valvepercent = 0
        value_multiplier = 0.001
        alt_index = 2: meter_index = 0
        lowhigh_select = 0
    End If

    ' Store the limit values of the gauge not being used (i.e., the connector that will be plugged
    ' into the calibrated meter
    real_vals(0) = FY1(lowhigh_select, alt_index)
    real_vals(1) = FY2(lowhigh_select, alt_index)
    real_vals(2) = FY1(lowhigh_select, alt_index + 1)
    real_vals(3) = FY2(lowhigh_select, alt_index + 1)
    
    ' Make the unused gauge an "ideal" flow meter of the correct range.
    FY2(1, alt_index) = InputBox(ts$(17), , "200000")       ' "Enter max. value of calibrated gauge (CCPM)."
    FY2(1, alt_index + 1) = 0.4 * FY2(1, alt_index)
    FY1(1, alt_index) = 0: FY1(1, alt_index + 1) = 0
    'If targetGauge$ = "XFM" Then
    '    fy2(1, alt_index) = 200000
    '    fy2(1, alt_index + 1) = 80000
    'ElseIf targetGauge$ = "HFM" Then
    '    fy2(1, alt_index) = 10000
    '    fy2(1, alt_index + 1) = 4000
    'End If
    
' 1) Figure out points to take (besides zero)
    While val(tempreturn$) < 4
        tempreturn$ = InputBox(ts$(3), , 4)     ' "Enter the number of readings to take (min. 4)"
        If tempreturn$ = "" Then Exit Sub
    Wend
    numreadings = val(tempreturn$)
    ReDim settingList(numreadings)
    ReDim results(numreadings, 2)
  '  If Check1.Value = vbChecked Then    ' Prompt for new max flow
        tempreturn$ = InputBox(ts$(18), , str$(maxrange))       ' "Enter a maximum flow value for the calibration (ccpm)"
        maxrange = val(tempreturn$)
  '  End If
    settingList(1) = maxrange * 0.4
    settingList(2) = -1                 ' indicates crossover from low range to high
    If numreadings = 4 Then
        settingList(3) = maxrange
    Else
        temprange = maxrange - (maxrange * 0.4) ' gives the distance from the crossover point to max pressure
        increment = temprange / (numreadings - 3)
        For i = 3 To numreadings - 1
            If settingList(i - 1) = -1 Then
                settingList(i) = settingList(i - 2) + increment
            Else
                settingList(i) = settingList(i - 1) + increment
            End If
        Next i
    End If
    
' 2) Set valves and zero regulator
    Dry_Chamber_Control "O"
    If ExtraPG Then Move_Valve 10, "C"     ' Protect LPG
    If targetGauge$ = "HFM" Then
        Move_Valve 9, "C"
    ElseIf targetGauge$ = "XFM" Then
        Move_Valve 9, "O"
    End If
    Zero_Reg
    Move_Valve 13, "C"
    Move_Valve 0, "C"
    
' 2a) Set values for readings
    If targetGauge = "LFM" Then
        channel = 0
    Else
        channel = 1                         ' high flow
    End If
    
    If targetGauge$ = "LFM" Or targetGauge$ = "HFM" Then
        u_low = 1: u_high = 0
        c_low = 3: c_high = 2
        v_cal = 1: v_uncal = 0
    Else
        u_low = 3: u_high = 2
        c_low = 1: c_high = 0
        v_cal = 0: v_uncal = 0
    End If
    
    If abort Then cal_aborted: Exit Sub
    
' 3) Take zero point
    Text2.Text = ts$(5)         ' "At zero point, waiting for stability"
    V2POS = cLimit
    Move2V2Pos
    waitseconds 10
    If abort Then cal_aborted: Exit Sub
    DoEvents
   ' tempreturn$ = ""
   ' While tempreturn$ = ""
   '     tempreturn$ = InputBox("Enter calibrated flow value:")
   ' Wend
   
    ' Uncalibrated
    u_avg = 0: c_avg = 0
    For k = 1 To 10
        HFLOW% = u_high: vflow% = v_uncal: lflow% = 0                    ' Set to low range
        ReadXReturnX4 channel
        u_avg = u_avg + x5
        ' results(1, 2) = Val(tempreturn$) * value_multiplier
    
        ' Now get the low range and adjust the zero for that signal separately
        HFLOW% = u_low: lflow% = 1
        ReadXReturnX4 channel
        c_avg = c_avg + x5  ' not really c_avg in this case, but oh well
   Next k
   
    zero_uncal = u_avg / 10
    c_avg = c_avg / 10
   
   ' correct the zero of the uncalibrated meter
    FY1(lowhigh_select, meter_index) = FY1(lowhigh_select, meter_index) - zero_uncal
    FY2(lowhigh_select, meter_index) = FY2(lowhigh_select, meter_index) - zero_uncal
    FY1(lowhigh_select, meter_index + 1) = FY1(lowhigh_select, meter_index + 1) - c_avg
    FY2(lowhigh_select, meter_index + 1) = FY2(lowhigh_select, meter_index + 1) - c_avg
    
    results(1, 1) = 0
    
    ' Now read the calibrated meter
    c_avg = 0
    For k = 1 To 10
        HFLOW% = c_low: vflow% = v_cal
        channel = 1         ' always one of the high-flow ranges
        ReadXReturnX4 channel
        c_avg = c_avg + x5
    Next k
    
    c_avg = c_avg / 10
    results(1, 2) = c_avg
    zero_cal = c_avg
    
    ' Print the results to screen
    statstring$ = ts$(6) + ":" + vbCrLf + Format$(results(1, 1), "#####0") + " cc/m " + ts$(7) + vbCrLf + Format$(results(1, 2), "#####0") + " cc/m " + ts$(8) + vbCrLf ' "Point taken"/"uncalibrated"/"calibrated"
    statstring$ = statstring$ + ts$(9) + " =" + Format$(Abs((results(1, 2) - results(1, 1)) / results(1, 2)) * 100, "###0.00") + "%" + vbCrLf       ' "Difference"
    Text2.Text = statstring$
    
    ' Open V1 or V2
    If targetGauge$ = "LFM" Then
        Move_Valve 0, "O"
    Else
        V2POS = cLimit + (oLimit - cLimit) * valvepercent / 100
        OpenV2Pos
    End If
        
    If abort Then cal_aborted: Exit Sub
    
' For each additional point:
    For i = 1 To numreadings - 1
    
        If abort Then cal_aborted: Exit Sub
        
        ' Reading the calibrated gauge to increase, so always on the hfm or xfm channel
        channel = 1
        at_target = False
        While Not at_target
        
            ' 4) Open V2 and regulator until pressure is achieved
            Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(19) + " ..." + "current = " + str$(x5)     ' "Reaching next flow value"
            inc_reg HScroll1.value
            
            ' 5) Wait for stability
            Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(11) + " ..."   ' "Waiting for stability"
            waitseconds 5
            
            Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(12) + " ..."   ' "Reading value"
            If settingList(i) = "-1" Then       ' At crossover point
                Pres% = u_high
                settingList(i) = settingList(i - 1)
            End If
            
            ReadXReturnX4 channel

            If x5 >= settingList(i) Then at_target = True
            If abort Then cal_aborted: Exit Sub
        
        Wend
        
        Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(20)        ' "Flow reached, waiting 10 seconds for stability"
        waitseconds 10
        
        ' 6) Prompt user for calibrated value
        'tempreturn$ = ""
        'While tempreturn$ = ""
        '    tempreturn$ = InputBox("Enter calibrated flow value:")
        'Wend
        
        ' 7) Store along with recorded value
        c_avg = 0: u_avg = 0
        For k = 1 To 10
            ' Uncal
            HFLOW% = u_high: vflow% = v_uncal
            channel = IIf(targetGauge$ = "LFM", 0, 1)
            ReadXReturnX4 channel
            u_avg = u_avg + x5
            'results(i, 2) = Val(tempreturn$) * value_multiplier
            ' If we're at the crossover point, read the high range at the same time
            If i = 1 Then
                HFLOW% = u_low
                ReadXReturnX4 channel
                crossover_reading = x5
            End If
                
            ' Now read calibrated value
            HFLOW% = c_high: vflow% = v_cal
            channel = 1
            ReadXReturnX4 channel
            c_avg = c_avg + x5
        Next k
        
        c_avg = c_avg / 10
        u_avg = u_avg / 10
        
        results(i, 2) = c_avg - zero_cal
        results(i, 1) = u_avg  ' - zero_uncal (we've adjusted the zero, so we shouldn't need to remove the original offset)

            
        
        
        statstring$ = statstring$ + vbCrLf + ts$(6) + ":" + vbCrLf + Format$(results(i, 1), "#####0") + " cc/m " + ts$(7) + vbCrLf + Format$(results(i, 2), "#####0") + " cc/m " + ts$(8) + vbCrLf  ' "Point taken"/"uncalibrated"/"calibrated"
        statstring$ = statstring$ + ts$(9) + " =" + Format$(Abs((results(i, 2) - results(i, 1)) / results(i, 2)) * 100, "###0.00") + "%" + vbCrLf       ' "Difference"
    Next i
    
    Text2.Text = statstring$ + vbCrLf + vbCrLf + ts$(21)        ' "Flow readings complete"
    
' 8) Eventually, do calculations and change values at end
    Zero_Reg
    V2POS = cLimit
    Move2V2Pos
    
    do_calcs
    
    ' Adjust the low-range offset so that the high and low ranges match at the
    ' crossover point. The crossover point value for the high range is stored in
    ' results(1,1); for the low range in crossover_reading
    FY1(lowhigh_select, meter_index + 1) = FY1(lowhigh_select, meter_index + 1) + (results(1, 1) - crossover_reading)
    FY2(lowhigh_select, meter_index + 1) = FY2(lowhigh_select, meter_index + 1) + (results(1, 1) - crossover_reading)
    
    ' Now work out the span correction for both high and low ranges
    span_multiplier = results(numreadings - 1, 2) - results(numreadings - 1, 1) ' cal-uncal
    span_multiplier = span_multiplier / results(numreadings - 1, 2)
  '  MsgBox (Str$(span_correct))
    ' if span_correct is positive, span has to be increased. Otherwise, decreased
    
    span_multiplier = 1 + span_multiplier
    
    ' Span is the difference between the zero point and max point
    span = FY2(lowhigh_select, meter_index) - FY1(lowhigh_select, meter_index)
    
    If Check1.value = vbUnchecked Then
        MsgBox (ts$(15) + " " + str$(span) + " " + ts$(16) + " " + str$(span * span_multiplier)) ' "Span will be changed from/"to"
    End If
    
    ' Reset the values for the "other" meter
    FY1(1, alt_index) = real_vals(0)
    FY2(1, alt_index) = real_vals(1)
    FY1(1, alt_index + 1) = real_vals(2)
    FY2(1, alt_index + 1) = real_vals(3)
    
    If Check1.value = vbUnchecked Then
        ' Set new span values here
        FY2(lowhigh_select, meter_index) = (span * span_multiplier) + FY1(lowhigh_select, meter_index)
        FY2(lowhigh_select, meter_index + 1) = (span * span_multiplier * 0.4) + FY1(lowhigh_select, meter_index + 1)
    
        ' Write the new values out
        WPPS "Capstuff", "FY1_" + Format(lowhigh_select) + Format(meter_index), Format(FY1(lowhigh_select, meter_index)), CSFile$
        WPPS "Capstuff", "FY2_" + Format(lowhigh_select) + Format(meter_index), Format(FY2(lowhigh_select, meter_index)), CSFile$
        WPPS "Capstuff", "FY1_" + Format(lowhigh_select) + Format(meter_index + 1), Format(FY1(lowhigh_select, meter_index + 1)), CSFile$
        WPPS "Capstuff", "FY2_" + Format(lowhigh_select) + Format(meter_index + 1), Format(FY2(lowhigh_select, meter_index + 1)), CSFile$
    End If

End Sub

Private Sub do_calcs()
' After a flow meter or PG run, report the changes that should be made to the zero and/or span

    Dim zero_difference As Single           ' Discrepancy at zero reading
    Dim max_difference As Single            ' Discrepancy at last reading
    Dim zero_correct As Integer             ' Correction factor for zero in counts
    Dim span_correct As Integer             ' Correction factor for span in counts
    Dim stext$
    Dim i As Integer
    Dim units$, direction$                  ' Formatting text
    
    cal_done = True
    
    If targetGauge$ = "HPG" Or targetGauge$ = "LPG" Then
        units$ = "PSI"
    Else
        units$ = "cc/m"
    End If
    
    zero_difference = results(1, 2) - results(1, 1)
    max_difference = results(numreadings - 1, 2) - results(numreadings - 1, 1)
    
    stext$ = Space$(10) + "U" + Space$(20) + "C" + Space$(18) + "%diff" + vbCrLf
    For i = 1 To 37
        stext$ = stext$ + "_"
    Next i
    stext$ = stext$ + vbCrLf + vbCrLf
    
    For i = 1 To numreadings - 1
        If units$ = "cc/m" Then
            stext$ = stext$ + Format$(results(i, 1), Space$(4) + "###000 cc/m" + Space$(6)) + Format$(results(i, 2), Space$(4) + "###000 cc/m" + Space$(6)) + Format$(((results(i, 2) - results(i, 1)) / results(i, 2)), Space$(4) + "#####0.00%" + Space$(7)) + vbCrLf
        Else
            stext$ = stext$ + Space$(4) + Format$(results(i, 1), "##0.#0 PSI") + Space$(5) + Format$(results(i, 2), "###0.#0 PSI") + Space$(5) + Format$((results(i, 2) - results(i, 1)) / results(i, 2), "##.00%") + vbCrLf
        End If
    Next i
    
    stext$ = stext$ + vbCrLf + vbCrLf
    stext$ = stext$ + ts$(22) + " " + Format$(zero_difference, "#####0.00 ") + units$ + "." + vbCrLf        ' "The zero points disagree by"
    stext$ = stext$ + ts$(23) + " " + Format$(max_difference, "#####0.00 ") + units$ + "." + vbCrLf + vbCrLf        ' "The max readings disagree by"
    
    If zero_difference = 0 Then
        stext$ = stext$ + ts$(24) + vbCrLf      ' "No zero correction is needed."
    Else
        If zero_difference < 0 Then             ' Uncalibrated is greater than calibrated
            direction$ = "down"
        Else
            direction$ = "up"
        End If
        stext$ = stext$ + "The zero point should be manually adjusted, or else "
        stext$ = stext$ + "the minimum and maximum " + units$ + " values in the capwin.ini file should be shifted "
        stext$ = stext$ + direction$ + " by " + Format$(Abs(zero_difference), "#####0.00 ") + units$ + "." + vbCrLf + vbCrLf
    End If
    
    If max_difference = 0 Then
        stext$ = stext$ + "No span correction is needed." + vbCrLf
    Else
        If max_difference < 0 Then
            direction = "down"
        Else
            direction = "up"
        End If
        stext$ = stext$ + "The span should be corrected by adjusting the maximum " + units$ + " values in the capwin.ini file "
        stext$ = stext$ + direction$ + " by " + Format$(Abs(max_difference), "#####0.00 ") + units$ + "." + vbCrLf
    End If
    
    Text2.Text = stext$
    Command5.Enabled = True
    
End Sub

Private Sub Command3_Click()
' Abort the current procedure

    abort = True
    ' If doing a flow calibration, v2 is open; but for pressure gauge, may be closed
    ' and need releasing
    If targetGauge$ = "LPG" Or targetGauge$ = "HPG" Then
        V2POS = V2POS = cLimit + (oLimit - cLimit) * 10 / 100
        OpenV2Pos
    End If
    Zero_Reg
    V2POS = cLimit
    Move2V2Pos
    cal_done = False
    Text2.Text = ts$(27)        ' "Calibration stopped by user."

End Sub

Private Sub Command4_Click()
' Adjust the zero point of the current transducer. If this is after a calibration is complete, we
' are setting the values based on the error. Otherwise, just set to the current value

    Dim zero_error As Single
    
    If Not cal_done Then            ' Simple zero - just adjust the initial value, don't worry about span
    End If
    

End Sub

Private Sub Command6_Click()
' Read all values from the transducers

    Dim fLL As Single, fLLc As Single
    Dim fLH As Single, fLHc As Single
    Dim fHL As Single, fHLc As Single
    Dim fHH As Single, fHHc As Single
    Dim fXL As Single, fXLc As Single
    Dim fXH As Single, fXHc As Single
    Dim pLL As Single, pLLc As Single
    Dim pLH As Single, pLHc As Single
    Dim pHL As Single, pHLc As Single
    Dim pHH As Single, pHHc As Single
    Dim temp$
    
    ' Low flow meter
    lflow% = 1
    ReadXReturnX4 0
    fLL = x5: fLLc = x4
    lflow% = 0
    ReadXReturnX4 0
    fLH = x5: fLHc = x4
    
    ' high flow meters
    HFLOW% = 1
    ReadXReturnX4 1
    fHL = x5: fHLc = x4
    HFLOW% = 0
    ReadXReturnX4 1
    fHH = x5: fHHc = x4
    vflow = 1
    HFLOW = 3
    ReadXReturnX4 1
    fXL = x5: fXLc = x4
    HFLOW = 2
    fXH = x5: fXHc = x4
    
    ' Pressure gauges
    Pres% = 3
    ReadXReturnX4 2
    pLL = x5: pLLc = x4
    Pres% = 2
    ReadXReturnX4 2
    pLH = x5: pLHc = x4
    Pres% = 1
    ReadXReturnX4 2
    pHL = x5: pHLc = x4
    Pres% = 0
    ReadXReturnX4 2
    pHH = x5: pHHc = x4
    
    temp$ = "Low flow meter: " + vbCrLf
    temp$ = temp$ + "Low range: " + str$(fLLc) + " counts, " + str$(fLL) + " cc/m" + vbCrLf
    temp$ = temp$ + "High range: " + str$(fLHc) + " counts, " + str$(fLH) + " cc/m" + vbCrLf
    temp$ = temp$ + "Med flow meter: " + vbCrLf
    temp$ = temp$ + "Low range: " + str$(fHLc) + " counts, " + str$(fHL) + " cc/m" + vbCrLf
    temp$ = temp$ + "High range: " + str$(fHHc) + " counts, " + str$(fHH) + " cc/m" + vbCrLf
    temp$ = temp$ + "High flow meter: " + vbCrLf
    temp$ = temp$ + "Low range: " + str$(fXLc) + " counts, " + str$(fXL) + " cc/m" + vbCrLf
    temp$ = temp$ + "High range: " + str$(fXHc) + " counts, " + str$(fXH) + " cc/m" + vbCrLf
    temp$ = temp$ + "Low pressure gauge: " + vbCrLf
    temp$ = temp$ + "Low range: " + str$(pLLc) + " counts, " + str$(pLL) + " PSI" + vbCrLf
    temp$ = temp$ + "High range: " + str$(pLHc) + " counts, " + str$(pLH) + " PSI" + vbCrLf
    temp$ = temp$ + "High perssure gauge: " + vbCrLf
    temp$ = temp$ + "Low range: " + str$(pHLc) + " counts, " + str$(pHL) + " PSI" + vbCrLf
    temp$ = temp$ + "High range: " + str$(pHHc) + " counts, " + str$(pHH) + " PSI" + vbCrLf
    
    MsgBox (temp$)
    
End Sub

Private Sub Form_Load()

    Dim count As Integer
    
        
    LoadTextStrings
    
    pressurestring$ = " PSI pressure gauge"
    flowstring$ = " cc/m flow meter"
    

    ' Fill the combo box
    With gaugelist
        .clear
        count = 0
        If FY2(0, 0) <> 0 Then
            .AddItem ts$(28)        ' "Low Flow"         ' (Str$(fy2(0, 0)) + flowstring$)
            count = count + 1
            readingIndex(count) = "LFM"
        End If
        If FY2(1, 0) <> 0 Then
            .AddItem ts$(29)        ' "Med. Flow"        ' (Str$(fy2(1, 0)) + flowstring$)
            count = count + 1
            readingIndex(count) = "HFM"
        End If
        If xhflow Then
            .AddItem ts$(30)        ' "High Flow"        ' (Str$(fy2(1, 2)) + flowstring$)
            count = count + 1
            readingIndex(count) = "XFM"
        End If
        If ExtraPG Then
            .AddItem ts$(31)        ' "Low Pressure"     ' (Str$(PY2(2)) + pressurestring$)
            count = count + 1
            readingIndex(count) = "LPG"
        End If
        If PY2(0) <> 0 Then
            .AddItem ts$(32)        ' "High Pressure"    ' (Str$(PY2(0)) + pressurestring$)
            count = count + 1
            readingIndex(count) = "HPG"
        End If

        
        .ListIndex = 1
    End With
    HScroll1.value = 100
    Label1.Caption = ts$(25) + ": " + str$(HScroll1.value)     ' "Regulator Step"
    
    Label2.Caption = ts$(26) + " " + gaugelist.List(gaugelist.ListIndex)        ' "Calibrating"
    
    abort = False
    Command5.Enabled = False                ' Change span
    Text2.Text = ""
    cal_done = False
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor

End Sub

Private Sub gaugelist_Click()

    Label2.Caption = ts$(26) + " " + gaugelist.List(gaugelist.ListIndex)    ' "Calibrating"
    targetGauge$ = readingIndex(gaugelist.ListIndex + 1)
    Command5.Enabled = False
    Text2.Text = ""
    cal_done = False
    If targetGauge$ = "LPG" Or targetGauge$ = "HPG" Then
        Command1.Enabled = True
        Command2.Enabled = False
    Else
        Command1.Enabled = False
        Command2.Enabled = True
    End If
    
End Sub

Private Sub HScroll1_Change()
    Label1.Caption = ts$(25) + ": " + str$(HScroll1.value)      ' "Regulator step"
End Sub

Private Sub Timer1_Timer()
' Update current readings
    
    Dim tempstring$
    Dim lowc As Long, lowv As Single
    Dim temp As Integer
    
    
    temp = vflow%
    tempstring$ = " cc/m"   ' Default
    Select Case targetGauge$
        Case "LPG"
            Pres = 0
            ReadXReturnX4 2
            lowc = x4: lowv = x5
            Pres = 2
            ReadXReturnX4 2
            tempstring$ = " PSI"
        Case "HPG"
            Pres = 2
            ReadXReturnX4 2
            lowc = x4: lowv = x5
            Pres = 0
            ReadXReturnX4 2
            tempstring$ = " PSI"
        Case "LFM"
            HFLOW = 2
            lflow = 1
            ReadXReturnX4 1
            lowc = x4: lowv = x5
            lflow = 0: vflow% = 1
            ReadXReturnX4 0
        Case "HFM"
            HFLOW = 0
            ReadXReturnX4 1
            lowc = x4: lowv = x5
            HFLOW = 2: vflow% = 1
            ReadXReturnX4 1
        Case "XFM"
            HFLOW = 2
            ReadXReturnX4 1
            lowc = x4: lowv = x5
            HFLOW = 0: vflow% = 0
            ReadXReturnX4 1
    End Select
        
    vflow% = temp
    lowreading.Caption = str$(lowc) + " counts, " + str$(lowv) + tempstring$
    highreading.Caption = str$(x4) + " counts, " + str$(x5) + tempstring$
    DoEvents

End Sub

Private Sub cal_aborted()
' User has pushed "stop" -- for now, just inform and exit

    Text2.Text = ts$(27)        ' "Calibration stopped by user."

End Sub

Public Sub LoadTextStrings()
' Load text elements for this form from external .ini file
    
    ' Form elements
    Autocal.Caption = get_thing("autocal", "window title", language$, Autocal.Caption, Autocal, default_font)
    set_fontstuff gaugelist, default_font
    Label2.Caption = get_thing("autocal", "label2", language$, Label2.Caption, Label2, default_font)
    Label5.Caption = get_thing("autocal", "label5", language$, Label5.Caption, Label5, default_font)
    Label6.Caption = get_thing("autocal", "label6", language$, Label6.Caption, Label6, default_font)
    Label7.Caption = get_thing("autocal", "label7", language$, Label7.Caption, Label7, default_font)
    Check1.Caption = get_thing("autocal", "check1", language$, Check1.Caption, Check1, default_font)
    set_fontstuff Label1, default_font
    set_fontstuff Text2, default_font
    Command1.Caption = gpps2("autocal", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    Command2.Caption = gpps2("autocal", "command2", language$, Command2.Caption)
    set_fontname Command2, default_font
    Command3.Caption = gpps2("autocal", "command3", language$, Command3.Caption)
    set_fontname Command3, default_font
    Command4.Caption = gpps2("autocal", "command4", language$, Command4.Caption)
    set_fontname Command4, default_font
    Command5.Caption = gpps2("autocal", "command5", language$, Command5.Caption)
    set_fontname Command5, default_font
    Command6.Caption = gpps2("autocal", "command6", language$, Command6.Caption)
    set_fontname Command6, default_font
    
    ' Other text
    ts$(1) = gpps2("autocal", "ts1", language$, "Enter min. value of calibrated gauge in PSI. (e.g. 14.7 for differential, 0 for absolute)")
    ts$(2) = gpps2("autocal", "ts2", language$, "Enter max. value of calibrated gauge in PSI.")
    ts$(3) = gpps2("autocal", "ts3", language$, "Enter the number of readings to take (min. 4)")
    ts$(4) = gpps2("autocal", "ts4", language$, "Enter a maximum pressure for the calibration")
    ts$(5) = gpps2("autocal", "ts5", language$, "At zero point, waiting for stability")
    ts$(6) = gpps2("autocal", "ts6", language$, "Point taken")
    ts$(7) = gpps2("autocal", "ts7", language$, "uncalibrated")
    ts$(8) = gpps2("autocal", "ts8", language$, "calibrated")
    ts$(9) = gpps2("autocal", "ts9", language$, "Difference")
    ts$(10) = gpps2("autocal", "ts10", language$, "Reaching next pressure")
    ts$(11) = gpps2("autocal", "ts11", language$, "Waiting for stability")
    ts$(12) = gpps2("autocal", "ts12", language$, "Taking value")
    ts$(13) = gpps2("autocal", "ts13", language$, "Pressure reached")
    ts$(14) = gpps2("autocal", "ts14", language$, "Pressure readings complete")
    ts$(15) = gpps2("autocal", "ts15", language$, "Span will be changed from")
    ts$(16) = gpps2("autocal", "ts16", language$, "to")
    ts$(17) = gpps2("autocal", "ts17", language$, "Enter max. value of calibrated gauge (CCPM).")
    ts$(18) = gpps2("autocal", "ts18", language$, "Enter a maximum flow value for the calibration (CCPM)")
    ts$(19) = gpps2("autocal", "ts19", language$, "Reaching next flow value")
    ts$(20) = gpps2("autocal", "ts20", language$, "Flow reached, waiting 10 seconds for stability")
    ts$(21) = gpps2("autocal", "ts21", language$, "Flow readings complete")
    ts$(22) = gpps2("autocal", "ts22", language$, "The zero points disagree by")
    ts$(23) = gpps2("autocal", "ts23", language$, "The max readings disagree by")
    ts$(24) = gpps2("autocal", "ts24", language$, "No zero correction is needed.")
    ts$(25) = gpps2("autocal", "ts25", language$, "Regulator step")
    ts$(26) = gpps2("autocal", "ts26", language$, "Calibrating")
    ts$(27) = gpps2("autocal", "ts27", language$, "Calibration stopped by user.")
    ts$(28) = gpps2("autocal", "ts28", language$, "Low Flow")
    ts$(29) = gpps2("autocal", "ts29", language$, "Med. Flow")
    ts$(30) = gpps2("autocal", "ts30", language$, "High Flow")
    ts$(31) = gpps2("autocal", "ts31", language$, "Low Pressure")
    ts$(32) = gpps2("autocal", "ts32", language$, "High Pressure")
    
End Sub
