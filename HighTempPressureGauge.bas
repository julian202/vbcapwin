Attribute VB_Name = "HighTempPressureGauge"
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

Private maxPressure As Integer
Public highTempCalibrationFilename As String
Private highTempCalibrations(19) As HighTempCalibrationInfo

'This function will open the calibration file and read all of the data into the
'highTempCalibrations data structure.
Public Sub readNewHighTempCalibrationFile()
    Dim filename As String
    Dim filenum As Long
    Dim zeroStr As String
    Dim twentyStr As String
    Dim fortyStr As String
    Dim sixtyStr As String
    Dim eightyStr As String
    Dim oneHundredStr As String
    Dim fieldCount As Long
    
    filename = EXE_Path$ + "\HighTempCalibration.txt"
    filenum = FreeFile
    
    Open filename For Input As #filenum
    Line Input #filenum, zeroStr
    Line Input #filenum, twentyStr
    Line Input #filenum, fortyStr
    Line Input #filenum, sixtyStr
    Line Input #filenum, eightyStr
    Line Input #filenum, oneHundredStr
    Close filenum
    
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

Public Function getNewHighTempPressure(ByVal count As Long, ByVal temperature As Long, ByVal Pres As Integer)
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
    Dim result As Single
    Dim range As Single
    
    If Pres = 0 Then
        maxPressure = 500
    Else
        maxPressure = 100
    End If
    
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

    If modValue = 0 Then
        Dim lowCount As Long
        Dim highCount As Long
        Dim lowReal As Single
        Dim highReal As Single
        
        'must be a temperature entry, no need to interpolate
        range = getPressureRange(count)
        lowCount = getLowCount(range, temperature)
        highCount = getHighCount(range, temperature)
        lowReal = getLowReal(range)
        highReal = getHighReal(range)
        
        Debug.Print "Range: " + Str$(range)
        Debug.Print "Low Count: " + Str$(lowCount)
        Debug.Print "High Count: " + Str$(highCount)
        Debug.Print "Low Real: " + Str$(lowReal)
        Debug.Print "High Real: " + Str$(highReal)
        
        result = GetValue(count, lowCount, highCount, lowReal, highReal)
    Else
        'must be a temperature in between entries, must interpolate.
        
        
        Dim lowCountBelow As Long
        Dim highCountBelow As Long
        Dim lowRealBelow As Single
        Dim highRealBelow As Single
        
        Dim lowCountAbove As Long
        Dim highCountAbove As Long
        Dim lowRealAbove As Single
        Dim highRealAbove As Single
        
        Dim resultBelow As Single
        Dim resultAbove As Single
        
        
        range = getPressureRange(count)
        lowCountBelow = getLowCount(range, (temperature - 10))
        highCountBelow = getHighCount(range, (temperature - 10))
        lowRealBelow = getLowReal(range)
        highRealBelow = getHighReal(range)
        
        resultBelow = GetValue(count, lowCountBelow, highCountBelow, lowRealBelow, highRealBelow)
        
        lowCountAbove = getLowCount(range, (temperature + 10))
        highCountAbove = getHighCount(range, (temperature + 10))
        lowRealAbove = lowRealBelow
        highRealAbove = highRealBelow
        
        resultAbove = GetValue(count, lowCountAbove, highCountAbove, lowRealAbove, highRealAbove)
        
        result = GetValue(temperature, temperature - 10, temperature + 10, resultBelow, resultAbove)
    End If
    Debug.Print "Result: " + Str(result)
    getNewHighTempPressure = result
End Function
Private Function GetValue(ByVal count As Long, ByVal lowCount As Long, ByVal highCount As Long, ByVal lowReal As Single, ByVal highReal As Single)
    Dim value As Single
    
    value = (((count - lowCount) * (highReal - lowReal)) / (highCount - lowCount)) + lowReal
    
    GetValue = value
End Function
Private Function getPressureRange(ByVal count As Long)
    Dim range As Single
    
    range = (count - 2000) / (62000 - 2000)
    getPressureRange = range
End Function
Private Function getCountOffSet(voltage As Single)
    Dim offset As Long
    
    offset = (voltage * 60000) / 10
    
    If Pres% = 1 Then
        offset = offset * 5
    End If
    getCountOffSet = offset
End Function
Private Function getLowCount(ByVal range As Single, ByVal temperature As Single) As Long
    Dim Index As Long
    Dim lowCount As Single
    
    Index = getTemperatureIndex(temperature)
    
    If range < (1 / 5) Then
        'lowCount = 2000 * (1 + (highTempCalibrations(Index).ZeroPercent))
        lowCount = 2000 + getCountOffSet(highTempCalibrations(Index).ZeroPercent)
    ElseIf range < (2 / 5) Then
        'lowCount = 14000 * (1 + (highTempCalibrations(Index).TwentyPercent - 2))
        lowCount = 14000 + getCountOffSet((highTempCalibrations(Index).TwentyPercent - 2))
    ElseIf range < (3 / 5) Then
        'lowCount = 26000 * (1 + (highTempCalibrations(Index).FortyPercent - 4))
        lowCount = 26000 + getCountOffSet((highTempCalibrations(Index).FortyPercent - 4))
    ElseIf range < (4 / 5) Then
        'lowCount = 38000 * (1 + (highTempCalibrations(Index).SixtyPercent - 6))
        lowCount = 38000 + getCountOffSet((highTempCalibrations(Index).SixtyPercent - 6))
    Else
        'lowCount = 50000 * (1 + (highTempCalibrations(Index).EightyPercent - 8))
        lowCount = 50000 + getCountOffSet((highTempCalibrations(Index).EightyPercent - 8))
    End If
    
    getLowCount = (lowCount)
End Function

Private Function getHighCount(ByVal range As Single, ByVal temperature As Single) As Long
    Dim Index As Long
    Dim highCount As Single
    
    Index = getTemperatureIndex(temperature)
    
    If range < (1 / 5) Then
        'highCount = 14000 * (1 + (highTempCalibrations(Index).TwentyPercent - 2))
        highCount = 14000 + getCountOffSet((highTempCalibrations(Index).TwentyPercent - 2))
    ElseIf range < (2 / 5) Then
        'highCount = 26000 * (1 + (highTempCalibrations(Index).FortyPercent - 4))
        highCount = 26000 + getCountOffSet((highTempCalibrations(Index).FortyPercent - 4))
    ElseIf range < (3 / 5) Then
        'highCount = 38000 * (1 + (highTempCalibrations(Index).SixtyPercent - 6))
        highCount = 38000 + getCountOffSet((highTempCalibrations(Index).SixtyPercent - 6))
    ElseIf range < (4 / 5) Then
        'highCount = 50000 * (1 + (highTempCalibrations(Index).EightyPercent - 8))
        highCount = 50000 + getCountOffSet((highTempCalibrations(Index).EightyPercent - 8))
    Else
        'highCount = 62000 * (1 + (highTempCalibrations(Index).OneHundredPercent - 10))
        highCount = 62000 + getCountOffSet((highTempCalibrations(Index).OneHundredPercent - 10))
    End If
    
    getHighCount = Int(highCount)
End Function

Private Function getLowReal(ByVal range As Single)
    Dim lowReal As Single
    
    If range <= 1 / 5 Then
        lowReal = 0
    ElseIf range <= 2 / 5 Then
        lowReal = maxPressure * 1 / 5
    ElseIf range <= 3 / 5 Then
        lowReal = maxPressure * 2 / 5
    ElseIf range <= 4 / 5 Then
        lowReal = maxPressure * 3 / 5
    Else
        lowReal = maxPressure * 4 / 5
    End If
    
    getLowReal = lowReal
    
End Function

Private Function getHighReal(ByVal range As Single)
    Dim highReal As Single
    
    If range <= 1 / 5 Then
        highReal = maxPressure * 1 / 5
    ElseIf range <= 2 / 5 Then
        highReal = maxPressure * 2 / 5
    ElseIf range <= 3 / 5 Then
        highReal = maxPressure * 3 / 5
    ElseIf range <= 4 / 5 Then
        highReal = maxPressure * 4 / 5
    Else
        highReal = maxPressure
    End If
    
    getHighReal = highReal
End Function

'This function will return the proper temperature index for the supplied temperature
Private Function getTemperatureIndex(ByVal temperature As Long) As Long
    Dim Index As Long
    
    Index = 1
    While highTempCalibrations(Index).temperature < temperature
        Index = Index + 1
    Wend
    getTemperatureIndex = Index
End Function
'This function should just strip the next leading data element from the string.
'Since the string is passed by reference, it should be left with that element being stripped
Private Function stripField(ByRef dataStr As String) As String
    Dim position As Long
    
    position = InStr(dataStr, Chr(9))
    stripField = Left(dataStr, position - 1)
    dataStr = Mid(dataStr, position + 1)
End Function
