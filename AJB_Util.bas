Attribute VB_Name = "AJB_Util"

Sub DisplayTestSetup()
    AaronTestDialog.Label1.Caption = "Current Unit: " + str$(current_unit%)
    AaronTestDialog.Label2.Caption = "Test Type: " + str$(TType%(current_unit%))
    AaronTestDialog.Label3.Caption = "Test Mode: " + str$(TMode%(current_unit%))
    AaronTestDialog.Show
End Sub
Sub switchToMv2()
    Dim desiredFlow As Single
    Dim mode As Integer
    Dim f1 As Single
    Dim f2 As Single
    Dim stableRange As Single

    'Record current pressure
    Dim endingPressure As Single
    Dim withinRange As Boolean

    switchingMVs = True

    ReadXReturnX4 1
    desiredFlow = x5
    
    If desiredFlow <= FY2(2, 0) Then
        stableRange = FY2(2, 0) * 0.001
    Else
        stableRange = FY2(2, 2) * 0.001
    End If

    x4 = raw_reading(2)
    endingPressure = (x4 - PX1(Pres%)) / (PX2(Pres%) - PX1(Pres%)) * (PY2(Pres%) - PY1(Pres%)) + PY1(Pres%)

    'record ending mv1 pos
    mv1EndPos = V2POS

    'record ending regpos
    mv1RegEndPos = REGPOS

    'close mv#1
    close_v2_completely
    
    'zero the regulator
    Zero_Reg
    
    'switch motor valve index
    motorValveIndex = 1
    HFLOW% = 0
    vflow% = 1
    using_hflow1 = True

    'reset v2 pos variable
    V2POS = 0

    'close v10a
    Move_Valve 9, "C"

    'reset open & close limit variables
    olimit = olimit2
    CLIMIT = CLIMIT2

    'get the desired flow rate
    'desiredFlow = FY2(1, 2)

    'set the needed pos of mv2
    V2POS = mv2_start_pos
    
    'open mv2
    OpenV2Pos
    
    'set the needed regulator pos
    REGPOS = mv2_reg_pos
    
    'increment the regulator to that pos
    inc_reg REGPOS
    
    'update the status window
    status.Label1.Caption = "Achieving Required Flow Rate: " + getFormattedFlow(desiredFlow)
    status.Label2.Caption = ""
    status.Label3.Caption = ""
    status.Command1.Visible = True
    
    status.Show
    
    V2FACTR = 50
    
    'achieve desired flow rate
    mode = 0
    x5 = 0
    withinRange = False
    Do Until x5 >= desiredFlow Or withinRange Or status.Command1.Enabled = False
        'read current flow
        ReadXReturnX4 1
        f1 = x5
        
        If mode = 0 Then
            'waiting for stable flow
            
            'wait half a second
            waitseconds 0.5
            
            'read current flow
            ReadXReturnX4 1
            f2 = x5
            
            'update label
            status.Label2.Caption = "Waiting for stable flow"
            status.Label3.Caption = "Current Flow: " + getFormattedFlow(f2)
            status.Refresh
            
            'wait for flow to stablize
            If (f2 - f1) < stableRange Then
                mode = 1
                status.Label2.Caption = "Stable Flow Found"
            Else
                x5 = 0
            End If
        ElseIf mode = 1 Then
            If V2POS < olimit Then
                V2POS = V2POS + V2INCR * V2FACTR * ver1or3
                If V2POS >= olimit Then
                    V2POS = olimit
                End If
                OpenV2Pos
                mode = 0
            Else
                'adjusting flow to desired flow
                If f2 < desiredFlow * 0.5 Then
                    inc_reg 20
                    mode = 0
                ElseIf f2 < desiredFlow * 0.75 Then
                    inc_reg 10
                    mode = 0
                ElseIf f2 < desiredFlow * 0.9 Then
                    inc_reg 5
                    mode = 0
                ElseIf f2 > desiredFlow * 1.5 Then
                    lower_reg 20
                    mode = 0
                ElseIf f2 > desiredFlow * 1.25 Then
                    lower_reg 10
                    mode = 0
                ElseIf f2 > desiredFlow * 1.1 Then
                    lower_reg 5
                    mode = 0
                Else
                    withinRange = True
                End If
            End If
        End If
    Loop
    
    switchingMVs = False
    
    status.Hide
    status.Command1.Visible = True
    
End Sub

Sub switchToMv3()
    Dim desiredFlow As Single
    Dim mode As Integer
    Dim endingPressure As Single
    Dim f1 As Single
    Dim f2 As Single
    Dim withinRange As Boolean
    Dim stableRange As Single
    
    switchingMVs = True
    
    ReadXReturnX4 1
    desiredFlow = x5
    
    If desiredFlow <= FY2(2, 0) Then
        stableRange = FY2(2, 0) * 0.001
    Else
        stableRange = FY2(2, 2) * 0.001
    End If
    
    x4 = raw_reading(2)
    endingPressure = (x4 - PX1(Pres%)) / (PX2(Pres%) - PX1(Pres%)) * (PY2(Pres%) - PY1(Pres%)) + PY1(Pres%)
    
    'record ending mv1 pos
    mv2EndPos = V2POS
    
    'record ending regpos
    mv2RegEndPos = REGPOS
    
    'close mv#1
    close_v2_completely
                
    'switch motor valve index
    motorValveIndex = 2
                    
    'reset v2 pos variable
    V2POS = 0
                
    'close v10a
    'Move_Valve 9, "C"
                
    'reset open & close limit variables
    olimit = olimit3
    CLIMIT = CLIMIT3
        
    'set the needed pos of mv3
    V2POS = mv3_start_pos
    
    'open mv2
    OpenV2Pos
    
    'set the needed regulator pos
    REGPOS = mv3_reg_pos
    
    'increment the regulator to that pos
    inc_reg REGPOS
    
    'update the status window
    status.Label1.Caption = "Achieving Required Flow Rate: " + getFormattedFlow(desiredFlow)
    status.Label2.Caption = ""
    status.Label3.Caption = ""
    status.Command1.Visible = True
    
    status.Show
    
    V2FACTR = 50
    
    'achieve desired flow rate
    mode = 0
    x5 = 0
    withinRange = False
    Do Until x5 >= desiredFlow Or withinRange Or status.Command1.Enabled = False
        'read current flow
        ReadXReturnX4 1
        f1 = x5
        
        If mode = 0 Then
            'waiting half a second for stable flow
            waitseconds 0.5
            
            'read current flow
            ReadXReturnX4 1
            f2 = x5
            
            'update label
            status.Label2.Caption = "Waiting for stable flow"
            status.Label3.Caption = "Current Flow: " + getFormattedFlow(f2)
            status.Refresh
            
            'wait for flow to stablize
            If (f2 - f1) < stableRange Then
                mode = 1
                status.Label2.Caption = "Stable Flow Found"
            Else
                x5 = 0
            End If
        ElseIf mode = 1 Then
            If V2POS < olimit Then
                V2POS = V2POS + V2INCR * V2FACTR * ver1or3
                If V2POS >= olimit Then
                    V2POS = olimit
                End If
                OpenV2Pos
                mode = 0
            Else
                'adjusting flow to desired flow
                If f2 < desiredFlow * 0.5 Then
                    inc_reg 20
                    mode = 0
                ElseIf f2 < desiredFlow * 0.75 Then
                    inc_reg 10
                    mode = 0
                ElseIf f2 < desiredFlow * 0.9 Then
                    inc_reg 5
                    mode = 0
                ElseIf f2 > desiredFlow * 1.5 Then
                    lower_reg 20
                    mode = 0
                ElseIf f2 > desiredFlow * 1.25 Then
                    lower_reg 10
                    mode = 0
                ElseIf f2 > desiredFlow * 1.1 Then
                    lower_reg 5
                    mode = 0
                Else
                    withinRange = True
                End If
            End If
        End If
    Loop
    
    switchingMVs = False
    
    status.Hide
    status.Command1.Visible = True
End Sub

Sub switchToMv1()

    switchingMVs = True
    
    'close mv2
    close_v2_completely
        
    'reset the motorvalveindex variable
    motorValveIndex = 0
    
    'reset valve limits
    SetV2Limits
    
    'Record current pressure
    Dim endingPressure As Single
    x4 = raw_reading(2)
    endingPressure = (x4 - PX1(Pres%)) / (PX2(Pres%) - PX1(Pres%)) * (PY2(Pres%) - PY1(Pres%)) + PY1(Pres%)
    
    'reset the v2pos variable
    V2POS = mv1EndPos
    
    'open mv1 to previous position
    OpenV2Pos
    
    'zero the regulator
    Zero_Reg
    
    'increment regulator to old position
    inc_reg mv1RegEndPos
    
    'update user what is happening
    status.Label1.Caption = "Switching from MV#2 to MV#1"
    status.Label2.Caption = "Setting Flow: " + str(motorValveSwitchFlow)
    status.Label3.Caption = ""
    status.Show
    
    'wait for a stable flow rate
    Dim stable As Boolean
    Dim f1 As Single
    Dim f2 As Single
    Dim t0 As Single
    Dim T As Single
    
    stable = False
    
    Do Until stable
        'read the flow
        ReadXReturnX4 1
        f1 = x5
        
        'reset t0
        t0 = Timer
        
        While T < 5
            T = Timer - t0
            ReadXReturnX4 1
            status.Label3.Caption = "Current Flow: " + getFormattedFlow(x5)
            status.Refresh
        Wend
        
        ReadXReturnX4 1
        f2 = x5
        
        If f2 - f1 < (FY2(1, 2) * 0.01) Then
            stable = True
        End If
    Loop
    
    'wait for the pressure to rebuild
    x5 = 0
    status.Label2.Caption = "Rebuilding pressure to: " + getFormattedPressure(endingPressure)
    While x5 < endingPressure
        ReadXReturnX4 2
        status.Label3.Caption = "Current Pressure: " + getFormattedPressure(x5)
        status.Refresh
        waitseconds 0.1
    Wend
    
    switchingMVs = False
End Sub
Sub setupSecondMotorValve(Flow As Single, endingPressure As Single)
    'AJB 12-21-09
    'Process to position motor valve to create comparable flow to first motor valve
    
    Dim currentFlow As Long
    Dim flowSet As Boolean
    Dim startTime As Single
    Dim currentTime As Single
    Dim flowStableCounter As Integer
    Dim flowInRange As Boolean
    
    flowStableCounter = 0
    flowSet = False
    status.Label1.Caption = "Switching motor valves to reach " + getFormattedFlow(Flow)
    status.Label2.Caption = "Setting second motor valve to reach current flow"
    status.Label3.Caption = ""
    status.Show
    Debug.Print "Status window should display"
    'Do an inital pulse
    Pulse_V2 0
    
    While Not flowSet
        'Read current flow
        ReadXReturnX4 1
        pulseFlow = (x5 + (FY2(2, 0) * 0.001))
        Debug.Print "PulseFlow: " + str(pulseFlow)
        If x5 < Flow And flowStableCounter = 0 Then
        'pulse motor valve
            Pulse_V2 0
        Else
            Pulse_V2 1
        End If
        
        'record start time
        startTime = Timer
        
        'while current flow is less than
        While x5 < pulseFlow And currentTime - startTime < 5
            currentTime = Timer
            
            ReadXReturnX4 1
            
            status.Label2.Caption = "Current Flow: " + getFormattedFlow(x5)
            status.Refresh
            
            waitseconds 1
        Wend
        
        If x5 > Flow Then
            flowStableCounter = flowStableCounter + 1
            If flowStableCounter = 5 Then
                flowSet = True
            End If
        End If
    Wend
    
    If x5 > (Flow + (FY2(2, 0) * 0.01)) Then
        status.Label1.Caption = "Correcting Flow to within " + getFormattedFlow((FY2(2, 0) * 0.01))
        flowSet = False
        Dim flow1 As Single
        Dim flow2 As Single
        
        While Not flowSet
            
            
            ReadXReturnX4 1
            flow1 = x5
                
            waitseconds 1
                
            ReadXReturnX4 1
            flow2 = x5
            
            If flow2 > Flow Then
                If flow2 > Flow + 10000 Then
                    lower_reg 10
                ElseIf flow2 > Flow + 5000 Then
                    lower_reg 5
                Else
                    lower_reg 1
                End If
            ElseIf flow2 < Flow Then
                If flow2 < Flow - 10000 Then
                    inc_reg 10
                ElseIf flow2 < Flow - 5000 Then
                    inc_reg 5
                Else
                    inc_reg 1
                End If
            End If
            
            'monitor the dropping flow
            While flow1 > flow2
                ReadXReturnX4 1
                flow1 = x5
                
                waitseconds 1
                
                ReadXReturnX4 1
                flow2 = x5
            Wend
            'wait for flow to stabilize
            
            If flow2 > (Flow - (FY2(2, 0) * 0.01)) And flow2 < (Flow + (FY2(2, 0) * 0.01)) Then
                flowSet = True
            End If
        Wend
    End If
    Debug.Print "Status Window should hide"
    
    status.Label1.Caption = "System waiting for pressure to return to ending pressure"
    status.Label2.Caption = "Ending Pressure: " + getFormattedPressure(endingPressure)
    status.Label3.Caption = ""
    
    ReadXReturnX4 2
    
    Do Until x5 >= endingPressure
        'read current pressure
        ReadXReturnX4 2
        status.Label3.Caption = "Current Pressure: " + getFormattedPressure(x5)
    Loop
    
    status.Hide
    status.Label1.Caption = ""
    status.Label2.Caption = ""
    status.Label3.Caption = ""
    
End Sub

Function getFormattedFlow(Flow As Single)

    Dim flowText As String
    
    flowText = FormatNumber(Flow, 2, , , True) + " SCCM"
    
    getFormattedFlow = flowText
    
End Function

Function getFormattedPressure(Pressure As Single)
    
    Dim pressureText As String
    
    pressureText = FormatNumber(Pressure, 3, , , True) + PU$
    
    getFormattedPressure = pressureText
    
End Function

Function turn_pump_on(Index As Integer)
    Dim speed As Byte
    speed = auto_wet_pump_speed
    If Index = 1 Then
        Send_RS232b "sA", speed
        Send_RS232 "MAF"
    Else
        Send_RS232b "sB", speed
        Send_RS232 "MBF"
    End If
End Function
Function turn_pump_off(Index As Integer)
    If Index = 1 Then
        Send_RS232b "sA", 0
        Send_RS232 "MAS"
    Else
        Send_RS232b "sB", 0
        Send_RS232 "MBS"
    End If
End Function
Function turn_pump_reverse(Index As Integer)
    Dim speed As Byte
    speed = auto_wet_pump_speed
    If Index = 1 Then
        Send_RS232b "sA", speed
        Send_RS232 "MAR"
    Else
        Send_RS232b "sB", speed
        Send_RS232 "MBR"
    End If
End Function
Function auto_wet_sample(valve As Integer)
    
    checkReserveTankLevel
    
    turn_pump_on 1
    
    'open wetting valve
    If wetting_valves_latch Then
        Move_Valve (valve - 1), "C"
    Else
        Move_Valve (valve - 1), "O"
    End If
    
    'wait for a number of seconds
    waitseconds auto_wet_wet_time
        
    'close the wetting valve
    If wetting_valves_latch Then
        Move_Valve (valve - 1), "O"
    Else
        Move_Valve (valve - 1), "C"
    End If
    
    turn_pump_reverse 1
    waitseconds auto_wet_reverse_time
    
    'turn off pump
    turn_pump_off 1
    
    Unload info_form
End Function

Sub checkReserveTankLevel()
    Dim tankLevel As Single
    
    If ReserveTankLevelChannel > 0 And ReserveTankFillLightValve >= 0 Then
        tankLevel = readReserveTankLevel
        If tankLevel < ReserveTankRefillPercent Then
            turnOnFillLight
        Else
            turnOffFillLight
        End If
    End If
End Sub

Sub turnOnFillLight()
    Move_Valve ReserveTankFillLightValve, "C"
End Sub

Sub turnOffFillLight()
    Move_Valve ReserveTankFillLightValve, "O"
End Sub
