Attribute VB_Name = "Watlows"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cpCopy As Long)

Global Const PROTOCOL_MODBUS = 0
Global Const MAX_COMMS_TIMOUT_TIME = 2000
Global Const MESSAGE_RECEIVED = 0

Public Function readWatlowViaModbus(channel%, getChar$, sendChar$) As String
    Dim message$, charRes$, retMessage$
    Dim register%, i%
        
    If athena = 0 Then
        ' must be watlow
        ' watlows use two different registers for their two channels
        register% = IIf((channel And 1) = 0, 360, 440)
        readWatlowViaModbus = performReadWatlowViaModbus(1, register%, getChar$, sendChar$)
    Else
        ' athena uses the same register, but different id numbers
        readWatlowViaModbus = performReadWatlowViaModbus((channel And 1) + 1, 8000, getChar$, sendChar$)
    End If
End Function

Public Function readWatlowSetPointViaModbus(channel%, getChar$, sendChar$) As String
    Dim message$, charRes$, retMessage$
    Dim register%, i%
        
    If athena = 0 Then
        ' must be watlow
        ' watlows use two different registers for their two channels
        register% = IIf((channel And 1) = 0, 2160, 2240)
        readWatlowSetPointViaModbus = performReadWatlowViaModbus(1, register%, getChar$, sendChar$)
    Else
        ' athena uses the same register, but different id numbers
        readWatlowSetPointViaModbus = performReadWatlowViaModbus((channel And 1) + 1, 8004, getChar$, sendChar$)
    End If
End Function

Public Function performReadWatlowViaModbus(ID%, register%, getChar$, sendChar$) As String
    Dim message$, retMessage$
    Dim i%
    Dim t0 As Double
    Dim td As Double
    Dim retCnt As Long
    Dim retval As Byte
    Dim ignore$
        
    message$ = build_modbus_read_message$(ID%, register%)

    'Clear any left over data from the serial port
    ' don't care what it is
    retCnt = 0
    Do
        ignore$ = RSEcho(getChar$, 1)
        If ignore$ = 0 Then
            retCnt = retCnt + 1
        Else
            retCnt = 0
        End If
    Loop Until (retCnt >= 5)

    ' send message out as fast as possible
    For i = 1 To LenB(message$)
        'Send_RS232 sendChar$ + Mid(message$, i, 1)
        RSOutput_Raw ChrB(Asc(sendChar$)) & MidB(message$, i, 1)
    Next i
    
    ' send a space character out the serial port normally
    ' This will have the effect of clearing out the port since we didn't
    ' wait for the echo of the above command to come back
    Send_RS232 " "
    
    retCnt = 0
    retMessage$ = ""
    t0 = Timer
    Do
        'charRes$ = RSEcho(getChar$, 1)
        retval = RSModbusGet(getChar$)
        retMessage$ = retMessage$ & ChrB$(retval)
        td = Timer
        If (td < t0) Then td = td + 86400
        retCnt = retCnt + 1
    Loop Until (((td - t0) > 10#) Or (retCnt >= 9))
    
    'Need to receive 9 characters back
    'For i = 1 To 9
    '    charRes$ = RSEcho(getChar$, 1)
    '    retMessage$ = retMessage$ & Chr$(charRes$)
    'Next i
    
    performReadWatlowViaModbus = convert_modbus_read_response(ID%, retMessage$)
End Function

Public Function RSModbusGet(ByVal getChar$) As Byte
Dim t0 As Double
Dim T As Double
Dim retval As Integer
' send getChar as normal character with echo
Send_RS232 getChar$
' get the extra character back in raw form
t0 = Timer
Do
    retval = RSInput_Raw()
    T = Timer
    If (T < t0) Then T = T + 86400
Loop Until ((T >= t0 + 2) Or (retval >= 0))
If retval < 0 Then
    RSModbusGet = 0
Else
    RSModbusGet = retval
End If
End Function

Public Sub setWatlowViaModbus(channel%, target As Single, getChar$, sendChar$)
    Dim message$, charRes$, retMessage$
    Dim register%, i%
    Dim curChar As Byte
    
    If athena = 0 Then
        ' must be watlow
        ' watlows use two different registers for their two channels
        register% = IIf((channel And 1) = 0, 2160, 2240)
        message$ = build_modbus_set_message$(1, register%, target)
    Else
        ' athena uses the same register, but different id numbers
        message$ = build_modbus_set_message$((channel And 1) + 1, 8004, target)
    End If
    
    For i = 1 To LenB(message$)
        'Send_RS232 sendChar$ + Mid(message$, i, 1)
        RSOutput_Raw ChrB(Asc(sendChar$)) & MidB(message$, i, 1)
    Next i
    
    ' send a space character out the serial port normally
    ' This will have the effect of clearing out the port since we didn't
    ' wait for the echo of the above command to come back
    Send_RS232 " "
    
    'Need to receive 8 characters back
    For i = 1 To 8
        charRes$ = RSEcho(getChar$, 1)
        retMessage$ = retMessage$ & Chr$(charRes$)
    Next i
    'ignore the return value - what could possibly go wrong?
End Sub

Public Function build_modbus_read_message$(ByVal ID%, ByVal registerr!)
    Dim register_hi!, register_lo!
    Dim reg_hi%, reg_lo%
    Dim message$

    register_hi! = registerr! \ 256
    register_lo! = registerr! - (register_hi! * 256)
    reg_hi% = register_hi!
    reg_lo% = register_lo!

    message$ = ChrB$(ID%) & ChrB$(3) & ChrB$(reg_hi%) & ChrB$(reg_lo%) & ChrB$(0) & ChrB$(2)

    build_modbus_read_message$ = message$ & calculate_crc_str(message$)
End Function

Public Function build_modbus_set_message$(ByVal ID%, ByVal registerr!, ByVal target As Single)
    Dim register_hi!, register_lo!
    Dim reg_hi%, reg_lo%
    Dim message$
    Dim valueBytes(4) As Byte
    Dim chr1, chr2 As Byte
    
    register_hi! = registerr! \ 256
    register_lo! = registerr! - (register_hi! * 256)
    reg_hi% = register_hi!
    reg_lo% = register_lo!
    CopyMemory valueBytes(0), target, 4
    
    message$ = ChrB$(ID%) & ChrB$(16) & ChrB$(reg_hi%) & ChrB$(reg_lo%) & ChrB$(0) & ChrB$(2) & ChrB$(4) & _
               ChrB$(valueBytes(1)) & ChrB$(valueBytes(0)) & ChrB$(valueBytes(3)) & ChrB$(valueBytes(2))
    
    build_modbus_set_message$ = message$ & calculate_crc_str(message$)
End Function

Public Function convert_modbus_read_response(ID%, message$) As String
    Dim temp_length%
    Dim char%
    Dim device_address%, command_type%, bytes_returned%
    Dim packet_chars$, crc_chars$, crc_conv_chars$
    Dim read_data_reg_bytes(4) As Byte

    temp_length% = LenB(message$)
    If temp_length% < 9 Then
        convert_modbus_read_response = -1
        Exit Function
    End If

    For char% = 0 To 3
        read_data_reg_bytes(char%) = 0
    Next char%

    device_address% = AscB(MidB$(message$, 1, 1))
    command_type% = AscB(MidB$(message$, 2, 1))
    bytes_returned% = AscB(MidB$(message$, 3, 1))

    If device_address% <> ID% Or command_type% <> 3 Or bytes_returned% <> 4 Then
        convert_modbus_read_response = -2
        Exit Function
    End If

    packet_chars$ = LeftB$(message$, (LenB(message$) - 2))
    crc_chars$ = RightB$(message$, 2)
    crc_conv_chars$ = calculate_crc_str(packet_chars$)

    If crc_chars$ <> crc_conv_chars$ Then
        convert_modbus_read_response = -3
        Exit Function
    End If

    read_data_reg_bytes(0) = AscB(MidB$(message$, 5, 1))
    read_data_reg_bytes(1) = AscB(MidB$(message$, 4, 1))
    read_data_reg_bytes(2) = AscB(MidB$(message$, 7, 1))
    read_data_reg_bytes(3) = AscB(MidB$(message$, 6, 1))

    convert_modbus_read_response = byteArrayToSingle(read_data_reg_bytes)
End Function

Public Function calculate_crc_str(message$) As String
    Dim crc&, char%, bit%, temp_crc_hi&, temp_crc_lo&
    Dim crc_hi%, crc_lo%

    crc& = 65535
    For char% = 1 To LenB(message$)
        crc& = crc& Xor AscB(MidB$(message$, char%, 1))
        For bit% = 1 To 8
            If (crc& And 1) <> 0 Then
                crc& = (crc& \ 2)
                crc& = (crc& Xor 40961)
            Else
                crc& = (crc& \ 2)
            End If
        Next bit%
    Next char%

    crc& = (crc& And 65535)

    temp_crc_hi& = (crc& \ 256)
    temp_crc_lo& = (crc& - (temp_crc_hi& * 256))

    crc_hi% = temp_crc_hi&
    crc_lo% = temp_crc_lo&
  
    calculate_crc_str = ChrB$(crc_lo%) + ChrB$(crc_hi%)
End Function

Public Function byteArrayToSingle(ByRef aByte() As Byte) As Single
    CopyMemory byteArrayToSingle, aByte(0), 4
End Function

