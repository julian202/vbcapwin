Attribute VB_Name = "HighTempCalibration"
Option Explicit

Public Type HighTempCalibrationInfo
    temperature As Long
    ZeroPercent As Single
    TwentyPercent As Single
    FortyPercent As Single
    SixtyPercent As Single
    EightyPercent As Single
    OneHundredPercent As Single
End Type

Private Const MAXPRESSURE = 500
Private Const PRESSURESTEP = 100

Public highTempCalibrationFilename As String
Private highTempCalibrations(19) As HighTempCalibrationInfo

'This function will open the calibration file and read all of the data into the
'highTempCalibrations data structure.
Public Sub readHighTempCalibrationFile()
    Dim fileName As String
    Dim fileNum As Long
    Dim zeroStr As String
    Dim twentyStr As String
    Dim fortyStr As String
    Dim sixtyStr As String
    Dim eightyStr As String
    Dim oneHundredStr As String
    Dim fieldCount As Long
    
    fileName = EXE_Path$ + "\HighTempCalibration.txt"
    fileNum = FreeFile
    
    Open fileName For Input As #fileNum
    Line Input #fileNum, zeroStr
    Line Input #fileNum, twentyStr
    Line Input #fileNum, fortyStr
    Line Input #fileNum, sixtyStr
    Line Input #fileNum, eightyStr
    Line Input #fileNum, oneHundredStr
    Close fileNum
    
    For fieldCount = 1 To 18
        With highTempCalibrations(fieldCount)
            .temperature = (fieldCount + 1) * 10
            .ZeroPercent = stripField(zeroStr)
            .TwentyPercent = stripField(twentyStr)
            .FortyPercent = stripField(fortyStr)
            .SixtyPercent = stripField(sixtyStr)
            .EightyPercent = stripField(eightyStr)
            .OneHundredPercent = stripField(oneHundredStr)
        End With
    Next
    
    With highTempCalibrations(19)
        .temperature = 200
        .ZeroPercent = zeroStr
        .TwentyPercent = twentyStr
        .FortyPercent = fortyStr
        .SixtyPercent = sixtyStr
        .EightyPercent = eightyStr
        .OneHundredPercent = oneHundredStr
    End With
End Sub

'The following function is the main calculation function.
'It will calculate and return the correct pressure
'value for the provided counts at the provided temperature.
Public Function getHighTempPressure(ByVal count As Long, ByVal temperature As Long)
    Dim modValue As Long
    Dim lowTemp As Long
    Dim highTemp As Long
    Dim tempDiff As Long
    Dim lowTempIndex As Long
    Dim highTempIndex As Long
    Dim lowTempLowVoltage As Single
    Dim lowTempHighVoltage As Single
    Dim highTempLowVoltage As Single
    Dim highTempHighVoltage As Single
    Dim lowTempLowPressure As Long
    Dim lowTempHighPressure As Long
    Dim highTempLowPressure As Long
    Dim highTempHighPressure As Long
    Dim voltage As Single
    Dim lowHighVoltageDiff As Single
    Dim lowCurVoltageDiff As Single
    Dim highHighVoltageDiff As Single
    Dim highCurVoltageDiff As Single
    Dim psiPerVolt As Single
    Dim lowPressure As Single
    Dim highPressure As Single
    Dim pressureDiff As Single
    Dim pressureDiffPerDegree As Single
    Dim result As Long
    
    'If the temperature is less than 20 we set it to 20, since this is our lower limit of acceptable
    'temperatures for this pressure gauge.  If the temperature is over 200 then we set it to 200
    'because that is the upper limit of acceptable temperatures for the pressure gauge.
    If temperature < 20 Then
        temperature = 20
    ElseIf temperature > 200 Then
        temperature = 200
    End If
    
    'The MOD function will return the remainder of dividing the temperature by 10.  This will
    'be used to determine if the supplied temperature is an entry in the calibration data structure.
    modValue = temperature Mod 10
    voltage = getVoltageFromCounts(count)
    
    If modValue = 0 Then    'Temperature is an entry
        'Retrieve the index in the data structure for the current temperature
        lowTempIndex = getTemperatureIndex(temperature)
        'Retrieve the temperature value.  This is probably not needed.
        lowTemp = getTemperature(lowTempIndex)
        
        'Retrieve the voltage in the data structure that is below the current voltage
        lowTempLowVoltage = getLowVoltage(lowTempIndex, voltage)
        'Retrieve the voltage in the data structure that is above the current voltage
        lowTempHighVoltage = getHighVoltage(lowTempIndex, voltage)
        'Retrieve the pressure in the data structure that correlates to the low voltage
        lowTempLowPressure = getLowPressure(lowTempIndex, voltage)
        'Retrieve the pressure in the data structure that correlates to the high voltage
        lowTempHighPressure = getHighPressure(lowTempIndex, voltage)
        'Determine the difference between the high and low voltages
        lowHighVoltageDiff = lowTempHighVoltage - lowTempLowVoltage
        'Determine the difference between the low and current voltages
        lowCurVoltageDiff = voltage - lowTempLowVoltage
        'Determine how many PSI should be added per differential voltage
        psiPerVolt = lowHighVoltageDiff / PRESSURESTEP
        'Calculate the intermediate pressure based off of the psiPerVolt calculation
        lowPressure = (lowCurVoltageDiff * psiPerVolt) + lowTempLowPressure
        
        result = lowPressure
    Else                    'Temperature is in between entries
        'Retrieve the index in the data structure for the temperature below the current one
        lowTempIndex = getTemperatureIndex(temperature - modValue)
        'Retrieve the index in the data structure for the temperature above the current one
        highTempIndex = getTemperatureIndex(temperature + (10 - modValue))
        'Retrieve the temperature value for the temperature below the current one
        lowTemp = getTemperature(lowTempIndex)
        'Retrieve the temperature value for the temperature above the current one
        highTemp = getTemperature(highTempIndex)
        
        'Perform all calculations for the low temperature just as if the modValue had been 0
        lowTempLowVoltage = getLowVoltage(lowTempIndex, voltage)
        lowTempHighVoltage = getHighVoltage(lowTempIndex, voltage)
        lowTempLowPressure = getLowPressure(lowTempIndex, voltage)
        lowTempHighPressure = getHighPressure(lowTempIndex, voltage)
        lowHighVoltageDiff = lowTempHighVoltage - lowTempLowVoltage
        lowCurVoltageDiff = voltage - lowTempLowVoltage
        psiPerVolt = lowHighVoltageDiff / PRESSURESTEP
        lowPressure = (lowCurVoltageDiff * psiPerVolt) + lowTempLowPressure
        
        'Perform all calculations for the high temperature just as if the modValue had been 0
        highTempLowVoltage = getLowVoltage(highTempIndex, voltage)
        highTempHighVoltage = getHighVoltage(highTempIndex, voltage)
        highTempLowPressure = getLowPressure(highTempIndex, voltage)
        highTempHighPressure = getHighPressure(highTempIndex, voltage)
        highHighVoltageDiff = highTempHighVoltage - highTempLowVoltage
        highCurVoltageDiff = voltage - highTempLowVoltage
        psiPerVolt = highHighVoltageDiff / PRESSURESTEP
        highPressure = (highCurVoltageDiff * psiPerVolt) + highTempLowPressure
        
        'Determine the difference between the calculated high and low pressures
        pressureDiff = highPressure - lowPressure
        'Determine how many PSI should be added per degree difference
        pressureDiffPerDegree = pressureDiff / 10
        'Determine how many degrees difference there are between the high and low temperature
        tempDiff = highTemp - lowTemp
        
        'Calculate the intermediate pressure between the high and low temperatures
        result = lowPressure + (pressureDiffPerDegree * tempDiff)
    End If
    
    getHighTempPressure = result
End Function

'This function will return the proper temperature index for the supplied temperature
Private Function getTemperatureIndex(temperature As Long) As Long
    Dim index As Long
    
    index = 1
    While highTempCalibrations(index).temperature <> temperature
        index = index + 1
    Wend
    getTemperatureIndex = index
End Function

'This function will return the proper temperature reading for the supplied temperature index.
Private Function getTemperature(index As Long) As Long
    getTemperature = (index + 1) * 10
End Function
'This function will return the lower voltage bound given the appropriate temperature index
'and the supplied current voltage reading.
Private Function getLowVoltage(index As Long, voltage As Single) As Single
    Dim result As Single
    
    With highTempCalibrations(index)
        If voltage < .ZeroPercent Then
            result = .ZeroPercent
        ElseIf voltage < .TwentyPercent Then
            result = .ZeroPercent
        ElseIf voltage < .FortyPercent Then
            result = .TwentyPercent
        ElseIf voltage < .SixtyPercent Then
            result = .FortyPercent
        ElseIf voltage < .EightyPercent Then
            result = .SixtyPercent
        ElseIf voltage < .OneHundredPercent Then
            result = .EightyPercent
        Else
            result = .OneHundredPercent
        End If
    End With
    
    getLowVoltage = result
End Function
'This function will return the lower pressure bound given the appropriate temperature index
'for the supplied voltage.
Private Function getLowPressure(index As Long, voltage As Single) As Single
    Dim result As Single
    
    With highTempCalibrations(index)
        If voltage < .ZeroPercent Then
            result = 0
        ElseIf voltage < .TwentyPercent Then
            result = 0
        ElseIf voltage < .FortyPercent Then
            result = MAXPRESSURE * 0.2
        ElseIf voltage < .SixtyPercent Then
            result = MAXPRESSURE * 0.4
        ElseIf voltage < .EightyPercent Then
            result = MAXPRESSURE * 0.6
        ElseIf voltage < .OneHundredPercent Then
            result = MAXPRESSURE * 0.8
        Else
            result = MAXPRESSURE
        End If
    End With
    
    getLowPressure = result
End Function

'This function will return the upper voltage bound given the appropriate temperature index
'and the supplied current voltage reading.
Private Function getHighVoltage(index As Long, voltage As Single) As Single
    Dim result As Single
    
    With highTempCalibrations(index)
        If voltage < .ZeroPercent Then
            result = .ZeroPercent
        ElseIf voltage <= .TwentyPercent Then
            result = .TwentyPercent
        ElseIf voltage <= .FortyPercent Then
            result = .FortyPercent
        ElseIf voltage <= .SixtyPercent Then
            result = .SixtyPercent
        ElseIf voltage <= .EightyPercent Then
            result = .EightyPercent
        ElseIf voltage <= .OneHundredPercent Then
            result = .OneHundredPercent
        Else
            result = .OneHundredPercent
        End If
    End With
    
    getHighVoltage = result
End Function

'This function will return the upper pressure bound given the appropriate temperature index
'for the supplied voltage.
Private Function getHighPressure(index As Long, voltage As Single) As Single
    Dim result As Single
    
    With highTempCalibrations(index)
        If voltage < .ZeroPercent Then
            result = 0
        ElseIf voltage <= .TwentyPercent Then
            result = MAXPRESSURE * 0.2
        ElseIf voltage <= .FortyPercent Then
            result = MAXPRESSURE * 0.4
        ElseIf voltage <= .SixtyPercent Then
            result = MAXPRESSURE * 0.6
        ElseIf voltage <= .EightyPercent Then
            result = MAXPRESSURE * 0.8
        ElseIf voltage <= .OneHundredPercent Then
            result = MAXPRESSURE
        Else
            result = MAXPRESSURE
        End If
    End With
    
    getHighPressure = result
End Function

'This function should take a count value that is read from the pressure gauge and
'return a related voltage value.  This will be used for conversions
Private Function getVoltageFromCounts(count As Long) As Single
    Dim val As Long
    Dim result As Single
    'Voltage limits are from 0 to 10
    'Count limits are from 2000 to 62000
    ' x/10 = (count - 2000) / 60000
    val = count - 2000
    result = (count / 60000) * 10
    
    getVoltageFromCounts = result
End Function

'This function should just strip the next leading data element from the string.
'Since the string is passed by reference, it should be left with that element being stripped
Private Function stripField(ByRef dataStr As String) As String
    Dim position As Long
    
    position = InStr(dataStr, Chr(9))
    stripField = Left(dataStr, position - 1)
    dataStr = Mid(dataStr, position + 1)
End Function
