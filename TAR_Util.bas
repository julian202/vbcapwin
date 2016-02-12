Attribute VB_Name = "TAR_Util"
'*****
' Demo
' OpVersion
' FUNCTION      ReadBalanceNotPenet
' PERPETRATOR   Tim Richards, 04 05 14
' DESCRIPTION   Read the balance, not the penetrometer, if we are using the
'               liquid extrusion porosimiter. In the INI file, this is the
'               case if the feature number 16 is added (liquid permeability)
'               and if the switch H2OPERM is set to "B". Search for code
'               inserted by this author.
' CALLED BY     ReadXReturnX4
' RETURNS       passes back ICounts (x4) and iScale (X5).
'               x5 is the actual scale value.
'               x4 is the count value.
'               Boolean True if successful, false if not
'
' iCounts and iScale are /not/ ByRef ---Byref added to other subs 040715 Weds TAR 11:40 AM
'
Function ReadBalanceNotPenet(iCounts As Long, iScale As Single) As Boolean
    ' copied from CAPSAT project CAPSUB.bas function RS232MT

    'Set up the scale to return correct iCounts (x4) values in response to iScale.
    'PEN500, PEN20500, P2PEN500, P2PEN20500
    'Pen20500 is low counts.
    'Pen500 is high counts.
    
' g_iMettler_fluid_min (grams)
' g_iMettler_fluid_max (grams)
' g_iMettler_fluid_density (grams/mL)
'
' (grams - grams) * mL/gram = mL = volume
'
' (g_iMettler_fluid_max-g_iMettler_fluid_min) / g_iMettler_fluid_density
' (200 - 0) / 1 == 200 mL
' water has a density of 1 g/mL. And one CC = 1 mL. So 1 gram of water is 1 mL of water.

    On Error GoTo timeout

    ReadBalanceNotPenet = True

'    If Demo = -1 Then
'        ReadBalanceNotPenet = False
'        Exit Function
'    End If

 ' Function RS232MT(outstr As String) As Single

    ' 1.3.6 changed timeout to 0.2 seconds and it now returns a -999 if it can't read
    ' the balance.  It is up to the reading program to determine if this is bad or
    ' not.  This routine no longer gives error messages or changes the running flag
    Dim timer_s As Single
    Dim i As Integer
    Dim errcnt As Integer
    Dim tempin As String
    Dim tempres As String
    Dim locoutstr As String
    Dim buffer As Variant
    Dim byttobuf(0) As Byte
    ' 1.3.6 current time to take better care of midnight switch
    Dim current_time As Single

    Dim temp As String      'TAR 050425

    errcnt = 0
top:
    tempin = TitleScrn.AuxComm.Input
    tempin = ""
    ' this is for future commands - right now, just gets weight.
    '   If Len(outstr) > 0 Then
    '       locoutstr=outstr
    '   Else
'        If OpVersion = 7.1 Then
            locoutstr = "SI" + vbCrLf
'        ElseIf OpVersion = 7.2 Then
'            locoutstr = Chr$(27) + "P"
'        End If
    '   endif
        timer_s = Timer
        TitleScrn.AuxComm.Output = locoutstr
        Do
            Do
                'DoEvents
                ' 1.3.6 begin
                current_time = Timer
                If current_time + 0.1 < timer_s Then current_time = current_time - 86400
                If current_time - timer_s > 0.2 Then
                    GoTo timeout
                End If
                ' 1.3.6 end
            Loop Until TitleScrn.AuxComm.InBufferCount > 0
            tempin = tempin + TitleScrn.AuxComm.Input
        Loop Until Right(tempin, 1) = vbLf
'        If OpVersion = 7.1 Then
            If Left$(tempin, 2) = "EL" Then
                ' power failure - turn the power back on
                TitleScrn.AuxComm.Output = "PWR 1" + vbCrLf
                ' wait 10 seconds
                timer_s = Timer
                While (Timer < timer_s + 10) And (Timer > timer_s - 1)
                    DoEvents
                Wend
                GoTo top
            End If
            tempres = Mid(tempin, 3, 1)
            If tempres <> "S" And tempres <> "D" Then GoTo comerr
            iScale = Val(Mid(tempin, 5)) - g_iMettler_fluid_min

            iScale = iScale - g_bBalanceNotPenet_ZeroPoint ' TAR 040809

' g_iMettler_fluid_min (grams)
' g_iMettler_fluid_max (grams)
' g_iMettler_fluid_density (grams/mL)
            iCounts = ( _
                    iScale / (g_iMettler_fluid_max - g_iMettler_fluid_min) _
                    * DAC_span _
                ) + DAC_zero

            If iCounts > DAC_over Then iCounts = DAC_over
            If iCounts < DAC_under Then iCounts = DAC_under

            iCounts = iCounts + g_iMettler_Negative_Counts_Offset 'TAR 040614
'DAC_OVER
'
'        ElseIf OpVersion = 7.2 Then
'            iScale = Val(tempin)
'        End If
    'End If
    Exit Function
errtest:
    errcnt = errcnt + 1
    Return
comerr:
    GoSub errtest
    If errcnt > 4 Then
        If RUNNING = -1 Then Exit Function
        MsgBox "The Mettler Balance is reporting off-scale."
        RUNNING = -1
'        Stop_Test
    Else
        GoTo top
    End If
    Exit Function
timeout:
'    GoSub errtest
'    If errcnt > 4 Then
'        If Running = -1 Then Exit Function
'        Temp = "Timeout error sending '" + locoutstr + "'"
'        Temp = Temp + vbCr + "The received response so far was '" + tempin + "'"
'        Temp = Temp + vbCr + "Cannot continue."
'        MsgBox Temp
'        Running = -1
'        Stop_Test
'    Else
'        GoTo top
'    End If
    ReadBalanceNotPenet = False
'End Function

End Function




Sub SendMettlerCommand(ByVal str_Command As String)
    Dim l_str_command As String
    Dim l_b_unstable As Boolean

    TAR_MsgForm.Caption = "Mettler Balance"
    TAR_MsgForm.Cancel.Visible = False
    TAR_MsgForm.lb.Visible = False
    TAR_MsgForm.OK.Visible = False
    TAR_MsgForm.Label1.Visible = False
    TAR_MsgForm.Show 0
    TAR_MsgForm.Refresh



    Dim timer_s As Single
    Dim i As Integer
    Dim errcnt As Integer
    Dim tempin As String
    Dim tempres As String
    Dim locoutstr As String
    Dim buffer As Variant
    Dim byttobuf(0) As Byte
    ' 1.3.6 current time to take better care of midnight switch
    Dim current_time As Single

    Dim temp As String      'TAR 050425


    errcnt = 0
    l_b_unstable = False

top:
    If l_b_unstable = True And TAR_MsgForm.OK.Visible = False Then
        Select Case str_Command
        Case "Zero Immediately"
            str_Command = "Zero"
        Case "Tare Immediately"
            str_Command = "Tare"
        Case Else
            GoTo send_mettler_exit
        End Select
    End If

    Select Case str_Command
    Case "Resetting Mettler Balance"
        l_str_command = "@"
    Case "Zero Immediately"
        l_str_command = "ZI"
    Case "Tare Immediately"
        l_str_command = "TI"
    Case "Zero"
        l_str_command = "Z"
    Case "Tare"
        l_str_command = "T"
    Case Else
        GoTo send_mettler_exit
    End Select


    TAR_MsgForm.Label.Caption = str_Command
    tempin = ""
    timer_s = Timer
    TitleScrn.AuxComm.Output = l_str_command + vbCrLf
    Do
        DoEvents
        Do
            ' 1.3.6 begin
            current_time = Timer
            If current_time + 0.1 < timer_s Then current_time = current_time - 86400
            If current_time - timer_s > 0.2 Then
                GoTo timeout
            End If
            ' 1.3.6 end
        Loop Until TitleScrn.AuxComm.InBufferCount > 0
        tempin = tempin + TitleScrn.AuxComm.Input
        Loop Until Right(tempin, 1) = vbLf
'        If OpVersion = 7.1 Then
        If Left$(tempin, 2) = "EL" Then
            ' power failure - turn the power back on
            TitleScrn.AuxComm.Output = "PWR 1" + vbCrLf
            ' wait 10 seconds
            timer_s = Timer
            While (Timer < timer_s + 10) And (Timer > timer_s - 1)
                DoEvents
            Wend
            GoTo top
        End If


    Select Case str_Command
    Case "Resetting Mettler Balance"
        If Left$(tempin, 5) <> "I4 A " Then GoTo top
    Case "Zero Immediately"
        If Left$(tempin, 3) <> "ZI " Then GoTo top
    Case "Tare Immediately"
        If Left$(tempin, 3) <> "TI " Then GoTo top
    Case "Zero"
        If Left$(tempin, 4) <> "Z A " Then GoTo top
    Case "Tare"
        If Left$(tempin, 4) <> "T S " Then GoTo unstable
    Case Else
        GoTo send_mettler_exit
    End Select

    GoTo send_mettler_exit
    
unstable:
    TAR_MsgForm.Label1.Visible = True
    TAR_MsgForm.Label1.Caption = "The balance is unstable. Trying again."
    TAR_MsgForm.OK.Visible = True
    TAR_MsgForm.OK.Caption = "Use Unstable"
    l_b_unstable = True
    GoTo top

errtest:
    errcnt = errcnt + 1
    Return

comerr:
timeout:
    GoSub errtest
    If errcnt > 4 Then
'        MsgBox "The Mettler Balance is reporting off-scale. Please check to ensure that it is stable."
        errcnt = 0
        GoTo top
    Else
        GoTo top
    End If

send_mettler_exit:
    Unload TAR_MsgForm
End Sub



' **********
' FUNCTION      Progress_Output_base
' RETURNS       String, the output base string
' PERPETRATOR   Tim Richards, Monday 6/21/04 9:25AM
' DESCRIPTION   outputs  ' "Pressure"/"Height"/"cm"/"PTarget" 'TAR040615
Function Progress_Output_Base(ByVal r_pressure_current As Single, ByVal r_height_or_mass As Single) As String
    Dim s_gm_cm As String
    Dim s_height_mass As String

    If g_bBalanceNotPenet = True Then
        s_gm_cm = capflow_ts(484)
        s_height_mass = capflow_ts(486)
    Else
        s_gm_cm = capflow_ts(485)
        s_height_mass = capflow_ts(243)
    End If

    Progress_Output_Base = _
        capflow_ts(229) + ": " + Xformat$(r_pressure_current, "###0.000 ") + PU$ + "    " + _
        s_height_mass + ": " + Xformat$(r_height_or_mass, "##0.0000 ") + s_gm_cm + "    "

End Function



' **********
' SUBROUTINE    Progress_Output_Time
' PERPETRATOR   Tim Richards, Monday 6/21/04 9:01AM
' DESCRIPTION   outputs  ' "Pressure"/"Height"/"cm"/"PTarget" 'TAR040615
Sub Progress_Output_Time(ByVal r_pressure_current As Single, ByVal r_height_or_mass As Single, ByVal r_time As Single)

    progress.Line25.Caption = Progress_Output_Base(r_pressure_current, r_height_or_mass) + _
        capflow_ts(74) + ": " + Xformat$(r_time, "###0.0") + " " + capflow_ts(245)
    progress.Line25.Refresh

End Sub



' **********
' SUBROUTINE    Progress_Output
' PERPETRATOR   Tim Richards, Monday 6/21/04 9:01AM
' DESCRIPTION   outputs  ' "Pressure"/"Height"/"cm"/"PTarget" 'TAR040615
Sub Progress_Output(ByVal r_pressure_current As Single, ByVal r_height_or_mass As Single, ByVal r_pressure_target As Single)

    progress.Line25.Caption = Progress_Output_Base(r_pressure_current, r_height_or_mass) + _
        capflow_ts(285) + ": " + Xformat$(r_pressure_target, "###0.000 ")
    progress.Line25.Refresh

End Sub



' **********
' SUBROUTINE    Progress_Output
' PERPETRATOR   Tim Richards, Monday 6/21/04 9:01AM
' DESCRIPTION   outputs  ' "Pressure"/"Height"/"cm"/"PTarget" 'TAR040615
Sub Elev_LqPerm_Wait(ByVal r_wait_time As Single, ByVal r_height_or_mass As Single, ByVal r_pressure_target As Single)
    Dim r_timeTarg As Single
    Dim r_countdown As Single

    'wait .25 second, let the pressure grow
    r_timeTarg = Timer + r_wait_time
    r_countdown = r_timeTarg - Timer
    While r_countdown > 0 And Not Aborted
        ReadXReturnX4 2
        Progress_Output (x5 - real_atm) * PCNV, r_height_or_mass, r_pressure_target
        progress.Line25.Caption = progress.Line25.Caption + "    Targeting"
        progress.Line25.Refresh
        DoEvents
        r_countdown = r_timeTarg - Timer
    Wend
End Sub



' **********
' SUBROUTINE    Progress_Output
' PERPETRATOR   Tim Richards, Monday 6/21/04 9:01AM
' DESCRIPTION   outputs  ' "Pressure"/"Height"/"cm"/"PTarget" 'TAR040615
Sub Elev_LqPerm_Settle(ByVal r_wait_time As Single, ByVal r_height_or_mass As Single)
    Dim r_timeTarg As Single
    Dim r_countdown As Single

    'wait .25 second, let the pressure grow
    r_timeTarg = Timer + r_wait_time
    r_countdown = r_timeTarg - Timer
    While r_countdown > 0 And Not Aborted
        ReadXReturnX4 2
        Progress_Output_Time (x5 - real_atm) * PCNV, r_height_or_mass, r_countdown
        progress.Line25.Caption = progress.Line25.Caption + "    Settling"
        progress.Line25.Refresh
        DoEvents
        r_countdown = r_timeTarg - Timer
    Wend
End Sub



' **********
' SUBROUTINE    Progress_Output
' PERPETRATOR   Tim Richards, Monday 6/21/04 9:01AM
' DESCRIPTION   outputs  ' "Pressure"/"Height"/"cm"/"PTarget" 'TAR040615
Sub Test_Done_Drain(ByVal r_target_mass As Single)
    Dim r_countdown As Single
    
    user_keypress = 0
    Do
        Move_Valve 12, "O"  'make sure valve 13 is open, even during the loop
        ReadXReturnX4 4
        r_countdown = x5 - r_target_mass
        ReadXReturnX4 2
        Progress_Output (x5 - real_atm) * PCNV, r_countdown, 0
        progress.Line25.Caption = progress.Line25.Caption + "    Draining (any key cancels)"
        progress.Line25.Refresh
        DoEvents
    Loop While r_countdown > 0 And Not Aborted And user_keypress = 0
End Sub



' **********
' SUBROUTINE    inc_dec_reg
' PERPETRATOR   Tim Richards, Tuesday 6/22/04 8:02AM
' DESCRIPTION   outputs  ' "Pressure"/"Height"/"cm"/"PTarget" 'TAR040615
Sub inc_dec_reg(ByVal i_plus_minus_counts As Integer)
    Dim i_counts As Integer

    i_counts = Abs(i_plus_minus_counts)
    If i_plus_minus_counts < 0 Then
        lower_reg i_counts
    Else
        inc_reg i_counts
    End If
End Sub



' **********
' FUNCTION      calc_counts_from_target_press
' PERPETRATOR   Tim Richards, Tuesday 6/22/04 8:07AM
' DESCRIPTION   outputs  ' "Pressure"/"Height"/"cm"/"PTarget" 'TAR040615
' RETURNS       counts as integer, r_wait_time as integer
'
' r_wait_time is /not/ ByVal --TAR 040715 12PM
'
Function calc_counts_from_target_press(ByVal r_curr_press As Single, ByVal r_pressure_target As Single, r_wait_time As Single) As Integer
    ' > 1   --> 200
    ' > .5  --> 50
    ' > .1  --> 10
    ' > .01 --> 2
    ' >     --> 1
    Dim r_pdiff As Single
    Dim i_sign As Integer

    r_pdiff = r_pressure_target - r_curr_press
    i_sign = r_pdiff / Abs(r_pdiff)  '1 or -1

    r_wait_time = 1.5
    If (Abs(r_pdiff) > 1) Then calc_counts_from_target_press = 200 * i_sign: Exit Function
    If (Abs(r_pdiff) > 0.5) Then calc_counts_from_target_press = 50 * i_sign: Exit Function

    r_wait_time = 0.75
    If (Abs(r_pdiff) > 0.1) Then calc_counts_from_target_press = 10 * i_sign: Exit Function

    r_wait_time = 0.5
    If (Abs(r_pdiff) > 0.01) Then calc_counts_from_target_press = 2 * i_sign: Exit Function

    r_wait_time = 0.5
    calc_counts_from_target_press = 1 * i_sign: Exit Function
End Function



' **********
' FUNCTION      reg_to_target_press
' PERPETRATOR   Tim Richards, Tuesday 6/22/04 8:15AM
' DESCRIPTION   outputs  ' "Pressure"/"Height"/"cm"/"PTarget" 'TAR040615
Sub reg_to_target_press(ByVal r_pressure_target As Single, ByVal r_current_mass_or_height As Single)
    Dim r_wait_time As Single
    Dim r_small_press As Single
    Dim b_retarget As Boolean

    'special case: target pressure = 0
    If r_pressure_target = 0 / PCNV + real_atm Then
        Zero_Reg
        Elev_LqPerm_Wait 1, r_current_mass_or_height, (r_pressure_target - real_atm) * PCNV
        Exit Sub
    End If
    
    'special case: target pressure < 2/3 of current pressure
    ReadXReturnX4 2
    If r_pressure_target < 2 / 3 * x5 Then
        Zero_Reg
        Elev_LqPerm_Wait 0.5, r_current_mass_or_height, (r_pressure_target - real_atm) * PCNV
    End If

    'special case: x5 pressure < .02, regulator is slow from 0.
    r_small_press = 0.02 / PCNV + real_atm 'inverse of (r_pressure_target - real_atm) * PCNV
    If r_pressure_target < r_small_press Then r_small_press = r_pressure_target
    ReadXReturnX4 2
    While x5 < r_small_press And Not Aborted
        If x5 < r_small_press / 5# Then inc_reg 100 Else inc_reg 50
        Elev_LqPerm_Wait 0.5, r_current_mass_or_height, (r_pressure_target - real_atm) * PCNV
        ReadXReturnX4 2
    Wend

    Do
        b_retarget = False
        ReadXReturnX4 2
        While Abs(r_pressure_target - x5) > 0.005 And Not Aborted      '040621
            inc_dec_reg calc_counts_from_target_press(x5, r_pressure_target, r_wait_time)
            Elev_LqPerm_Wait r_wait_time, r_current_mass_or_height, (r_pressure_target - real_atm) * PCNV
            ReadXReturnX4 2
            b_retarget = True
        Wend
        
        ' at this point we are very close to the target pressure. Wait one second for it to settle.
        Elev_LqPerm_Wait 1, r_current_mass_or_height, (r_pressure_target - real_atm) * PCNV
    Loop While b_retarget = True And Not Aborted

End Sub



'' **********
'' FUNCTION      Find_Machine_Port_RS232
'' PERPETRATOR   Tim Richards, Monday 7/19/04 8:51AM
'' DESCRIPTION   Finds the PMI machine on RS232 if there is one and passes back its port number.
'' RETURNS       Boolean, True if port found, false if not found
'' PASSES BACK   passback_PortNum as Integer, the port number that the machine is attached to
'Function Find_Machine_Port_RS232(passback_iPortNum As Integer) As Boolean
'    ' At the time of this writing there are 6 ports on TitleScrn (CAPMAIN.FRM)
'    Dim i As Integer
'    Const iNumPorts = 6
'
'    For i = 1 To iNumPorts
'
'    Next i
'End Function



' **********
' SUBROUTINE    Init_Port_RS232_Base
' PERPETRATOR   Tim Richards, Monday 7/19/04 9:50AM
' DESCRIPTION   Sets up the port requested.
' silent subroutine
'Function Init_Port_RS232_Base()
'' we can't make changes to the port settings while it is closed.
'' close it and then reopen it after the settings are changed.
'    If TitleScrn.FrmCtrl_MSComm_RS232(iPortNum).PortOpen Then
'        TitleScrn.FrmCtrl_MSComm_RS232(iPortNum).PortOpen = False
'    End If
'
'End Function



'' **********
'' SUBROUTINE    Init_Port_RS232_Machine
'' PERPETRATOR   Tim Richards, Monday 7/19/04 9:50AM
'' DESCRIPTION   Sets up the port requested.
'' silent subroutine
'Sub Init_Port_RS232_Machine()
'
'End Sub



'' **********
'' SUBROUTINE    Init_Port_RS232_Mettler
'' PERPETRATOR   Tim Richards, Monday 7/19/04 9:50AM
'' DESCRIPTION   Sets up the port requested.
'' silent subroutine
'Sub Init_Port_RS232_Mettler()
'
'End Sub
'
