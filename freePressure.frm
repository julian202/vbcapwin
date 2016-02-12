VERSION 5.00
Begin VB.Form freePressure 
   BackColor       =   &H000000FF&
   Caption         =   "Free-pressure permeability test"
   ClientHeight    =   9375
   ClientLeft      =   2520
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   5805
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame1 
         Caption         =   "Setup"
         Height          =   6615
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   3495
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2520
            TabIndex        =   18
            Text            =   "2000"
            Top             =   4200
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   2520
            TabIndex        =   20
            Text            =   "30"
            Top             =   4680
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2520
            TabIndex        =   16
            Top             =   3720
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   2520
            TabIndex        =   22
            Text            =   "30"
            Top             =   5160
            Width           =   735
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   2520
            TabIndex        =   14
            Top             =   3240
            Width           =   735
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   2520
            TabIndex        =   12
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   1440
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   1440
            TabIndex        =   4
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   1440
            TabIndex        =   6
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   2520
            TabIndex        =   10
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   2520
            TabIndex        =   24
            Text            =   "10"
            Top             =   5520
            Width           =   735
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   2520
            TabIndex        =   26
            Text            =   "4"
            Top             =   6000
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Number of points per sample"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   3720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Regulator setting (0 - 4000)"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   4200
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Valve setting (0 - 100%)"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   19
            Top             =   4680
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stability time (seconds)"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   21
            Top             =   5160
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Number of samples"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   3240
            Width           =   2055
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   3240
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nozzle OD (cm)"
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   11
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Product group"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Sample OD (in)"
            Height          =   255
            Index           =   7
            Left            =   720
            TabIndex        =   7
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Operator"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Sample ID"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Thickness (in)"
            Height          =   255
            Index           =   10
            Left            =   720
            TabIndex        =   9
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   3240
            Y1              =   5880
            Y2              =   5880
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Pause between samples (sec)"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   23
            Top             =   5520
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Number of samples per row in report spreadsheet"
            Height          =   495
            Index           =   12
            Left            =   120
            TabIndex        =   25
            Top             =   6000
            Width           =   2295
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start test"
         Height          =   495
         Left            =   3840
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3840
         TabIndex        =   29
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Take Reading (space bar)"
         Height          =   1335
         Left            =   3840
         TabIndex        =   30
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         Height          =   975
         Left            =   240
         TabIndex        =   31
         Top             =   7680
         Width           =   4935
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   615
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "C"
         Height          =   375
         Left            =   4680
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "Data file (click ""c"" button to change):"
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
         Left            =   360
         TabIndex        =   34
         Top             =   240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "freePressure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pathname$
Dim STn As Integer
Dim STv As Single
Dim STre As Integer
Dim STfn As String
Dim take_reading As Boolean         ' time to take a reading


Private Sub Command1_Click()
' Begin the test

    Dim i As Integer, j As Integer, k As Integer
    Dim numPoints As Integer        ' Number of data points to take per sample
    Dim regSetting As Integer       ' Regulator setting (0 - 4000 counts)
    Dim valvepercent As Single      ' V2 percentage
    Dim errorString$
    Dim ti As Single
    Dim fn As Integer
    Dim sampleDiam As Single        ' Sample OD in inches
    Dim nozzleDiam As Single        ' Nozzle OD in inches
    Dim numSamples As Integer       ' Number of samples being tested
    Dim sampleThick As Single       ' sample thickness in inches
    Dim startTime As Long         ' timer
    Dim stabilityTime As Single     ' how long to wait for stability
    Dim statusText$
    Dim temp$
    Dim flowValues() As Single, pressValues() As Single
    Dim samplePauseTime As Integer  ' Number of seconds to wait between each sample
    ' For excel file output
    Dim ExcelApp As Object
    Dim currentLine As Integer
    Dim samplesPerRow As Integer    ' Number of samples to list in a row in the excel sheet
    Dim samplesPrinted As Integer   ' Number of samples printed so far
    Dim sampleMinRange As Integer, sampleMaxRange As Integer    ' Range of samples to print to a specific row
    Dim partialRow As Boolean       ' Flag for a partial (last) row in a report
    Dim lastRow As Boolean          ' Flag for the last row
    Dim max As Integer
    
    ' First validate input
    numPoints = val(Text1.Text)
    regSetting = val(Text2.Text)
    valvepercent = val(Text3.Text)
    stabilityTime = val(Text4.Text)
    samplePauseTime = val(Text12.Text)
    errorString$ = ""
    
    If numPoints < 1 Then errorString$ = errorString$ + vbCrLf + "The number of points must be greater than 0."
    If regSetting < 0 Or regSetting > 4000 Then errorString$ = errorString$ + vbCrLf + "The regulator setting must be between 0 and 4000."
    If valvepercent < 0 Or valvepercent > 100 Then errorString$ = errorString$ + vbCrLf + "The valve setting must be a percentage between 0 and 100."
    If pathname$ = "" Then errorString$ = errorString$ + vbCrLf + "The data filename is invalid."
    If samplePauseTime < 0 Then samplePauseTime = 0
    If stabilityTime < 0 Then
        stabilityTime = 0
        Text4.Text = "0"
    End If
    
    ' If there's a problem, let the user know and abort the test.
    If errorString$ <> "" Then
        MsgBox ("You have entered invalid values in your test setup." + vbCrLf + errorString$)
        Exit Sub
    End If
    
    ' If we're still here, we can start the test
    DoEvents
    
    sampleDiam = myVal(Text8.Text)
    If sampleDiam <= 0 Then sampleDiam = 0.0001
    nozzleDiam = myVal(Text6.Text)
    If nozzleDiam <= 0 Then nozzleDiam = 0.0001
    numSamples = myVal(Text5.Text)
    If numSamples < 1 Then numSamples = 1
    sampleThick = myVal(Text11.Text)
    If sampleThick <= 0 Then sampleThick = 0.0001
    
    
    
    ' Set regulator and valve
    STn = numPoints
    STv = valvepercent
    STre = regSetting
    STfn = pathname$
    
    Command1.Enabled = False
    Command2.Enabled = True
    
    Zero_Reg
    
    ' Get amospheric pressure for later
    ReadXReturnX4 2
  '  p_atm = x5
    
    inc_reg STre
    V2INCR = STv
    V2POS = CLIMIT + (olimit - CLIMIT) * V2INCR / 100
    OpenV2Pos
    
    ReDim flowValues(numSamples, numPoints)
    ReDim pressValues(numSamples, numPoints)

    Command3.Enabled = True
    Command3.SetFocus
    ' Two loops now: number of samples and number of points
    Label2.Caption = "Test running ..."
    
    For i = 1 To numSamples

        For j = 1 To numPoints
            
            statusText$ = "Sample" + Str$(i) + ":  Point" + Str$(j) + " of" + Str$(numPoints)
            Label2.Caption = statusText$
            take_reading = False
            ' User can either hit the space bar or "take point" button to record, or else wait for
            ' the stability time to time out.
            startTime = Timer
            statusText$ = statusText$ + vbCrLf + "Waiting for stability ..."
            Label2.Caption = statusText$
            While ((Timer - startTime) <= stabilityTime) And Not take_reading
                DoEvents
            Wend
            
            ' Time to take a reading! Either we've timed out on the stability or user has
            ' forced an early reading
            ReadXReturnX4 1
            flowValues(i, j) = x5
            ReadXReturnX4 2
            pressValues(i, j) = x5
            startTime = Timer
            If j < numPoints Then       ' Don't do this for the last point because we've got another delay right after it.
                While ((Timer - startTime) <= 5)
                    DoEvents
                    Label2.Caption = "Point taken. Waiting 5 seconds before continuing .... " + Str$(Int(Timer - startTime))
                Wend
            End If
        Next j
        If i < numSamples Then          ' Don't do this for the last sample
            startTime = Timer
            While ((Timer - startTime) <= samplePauseTime)
                DoEvents
                Label2.Caption = "Sample " + Str$(i) + " complete. Change sample now -- waiting" + Str$(samplePauseTime) + " seconds before continuing .... " + Str$(Int(Timer - startTime))
            Wend
        End If
    Next i
            
 '   While pointcount < STn
 '       Label2.Caption = "Test running ... Ready"
 '       DoEvents
 '   Wend
    
  '  abc = 3.14159265 * ((nozzleDiam / 2) / 2.54) ^ 2
    
    samplesPerRow = val(Text13.Text)
    If samplesPerRow < 1 Then samplesPerRow = 1

    Set ExcelApp = CreateObject("excel.application")
    If Err = 0 Then
        With ExcelApp
            .sheetsinnewworkbook = 1
            .workbooks.Add
            currentLine = 1
            .cells(currentLine, 5) = "PERMEABILITY TEST"
            currentLine = currentLine + 2
            .cells(currentLine, 1) = "Operator:"
            .cells(currentLine, 3) = Text9.Text
            currentLine = currentLine + 1
            .cells(currentLine, 1) = "Date:"
            .cells(currentLine, 3) = Format$(Now, "MM/dd/yyyy")
            currentLine = currentLine + 2
            .cells(currentLine, 1) = "Product Group:"
            .cells(currentLine, 3) = Text7.Text
            currentLine = currentLine + 1
            .cells(currentLine, 1) = "Sample ID:"
            .cells(currentLine, 3) = Text10.Text
            For i = 1 To myVal(Text5.Text)
                currentLine = currentLine + 1
                .cells(currentLine, 3) = "Wheel" + Str$(i)
            Next i
            currentLine = currentLine + 1
            .cells(currentLine, 1) = "Nozzle OD (cm):"
            .cells(currentLine, 3) = Str$(nozzleDiam)
            currentLine = currentLine + 1
            .cells(currentLine, 1) = "Pressure (PSI):"
            .cells(currentLine, 3) = pressValues(1, 1) * 0.00689475729
            currentLine = currentLine + 2
            .cells(currentLine, 1) = "Number of data points:"
            .cells(currentLine, 3) = Str$(numPoints)
            currentLine = currentLine + 1
            .cells(currentLine, 1) = "Sample OD (in):"
            .cells(currentLine, 3) = Str$(sampleDiam)
            currentLine = currentLine + 1
            .cells(currentLine, 1) = "Sample thickness (in):"
            .cells(currentLine, 3) = Str$(sampleThick)
            currentLine = currentLine + 2
            .cells(currentLine, 1) = "SUMMARY"
            currentLine = currentLine + 2
            .cells(currentLine, 4) = Text10.Text
            currentLine = currentLine + 1
            
            samplesPrinted = 0
            partialRow = False: lastRow = False
                          
            For i = 1 To numSamples / samplesPerRow + 1         ' Number of row sets we have to print
                
                ' For each row, figure out the range of samples we'll be printing
                sampleMinRange = samplesPrinted + 1
                sampleMaxRange = IIf(samplesPrinted + samplesPerRow < numSamples, sampleMinRange + samplesPerRow - 1, numSamples)
               ' If sampleMaxRange - sampleMinRange + samplesPrinted > numSamples Then sampleMaxRange = numSamples - samplesPrinted
                
                If samplesPerRow + samplesPrinted > numSamples Then samplesPerRow = numSamples - samplesPrinted
                
                ' Print the headers for the row
                For j = 1 To samplesPerRow
                    .cells(currentLine, 1) = "Data Pt."
                    .cells(currentLine, j * 4) = "WHL" + Str$(samplesPrinted + j)
                    .cells(currentLine + 1, (j * 4) - 1) = "Flow (cc)"
                    .cells(currentLine + 1, (j * 4)) = "Pres (MPa)"
                    .cells(currentLine + 1, (j * 4) + 1) = "Perm (cc/Mpa)"
                    .cells(currentLine + 1, (j * 4) - 1).Columns.autofit
                    .cells(currentLine + 1, (j * 4) + 1).Columns.autofit
                    .cells(currentLine + 1, (j * 4)).Columns.autofit
                Next j
                
                currentLine = currentLine + 1
                
                ' Print all the data points for the row
                For j = 1 To numPoints          ' Numpoints is the same for all samples
                    currentLine = currentLine + 1
                    If samplesPrinted < numSamples Then .cells(currentLine, 1) = Str$(j)
                    For k = sampleMinRange To sampleMaxRange
                        .cells(currentLine, ((k - sampleMinRange + 1) * 4) - 1) = Format$(flowValues(k, j), "0")
                        .cells(currentLine, ((k - sampleMinRange + 1) * 4)) = Format$(pressValues(k, j) * 0.00689475729, "0.000E-##") ' convert to MPa
                        .cells(currentLine, ((k - sampleMinRange + 1) * 4) + 1) = Format$(flowValues(k, j) / (pressValues(k, j) * 0.00689475729), "####0.0###")
                    Next k
                Next j
                
                ' Increment our counter
                samplesPrinted = samplesPrinted + samplesPerRow
                
                ' Kick down a couple of lines
                currentLine = currentLine + 2
                
            Next i          ' End of row printing
            
            
            .activeworkbook.Saveas pathname$
            
        End With
        
        On Error Resume Next
        ExcelApp.Quit
        Set ExcelApp = Nothing
        
    End If
    
    Command1.Enabled = True
    Command2.Enabled = False
       
    ' Finish up
    Zero_Reg
    ' Set regulator and valve
    V2POS = CLIMIT
    Move2V2Pos
    
    Label2.Caption = "Test done. Please remove sample."
    
End Sub

Private Sub Command2_Click()
' Cancel button: stop the test and exit

    If MsgBox("Do you want to abort the current test?", vbYesNo) = vbNo Then Exit Sub
    Zero_Reg
    ' Set regulator and valve
    V2POS = CLIMIT
    Move2V2Pos
    Label2.Caption = "Test stopped by user."
        
End Sub

Private Sub Command3_Click()
' Read values

   
  '  Label2.Caption = "Test running ... Waiting" + Str$(STwt) + " seconds"
    'T = Timer
    'While (Timer - T) < STwt: DoEvents: Wend
  '  waitseconds STwt
    
  '  pointcount = pointcount + 1
    take_reading = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Randomize
    Select Case KeyCode
        Case vbKeySpace
            ' If user presses the space bar, it's the same as a "take value" click
          '  pointcount = pointcount + 1
            take_reading = True
        Case Else
     End Select

End Sub


Private Sub Command4_Click()
' Change the current filename

    ' Configure and open file selection box
    fsel_path$ = EXE_Path$ + "\data\*.xls"
    fsel_title$ = "Choose output filename"
    fsel_name$ = ""
    fsel_io = False
    fsel Me.hwnd
   
    If fsel_return = "" Then Exit Sub
    
    pathname$ = fsel_return
    Label3.Caption = pathname$
    
End Sub

Private Sub Form_Load()

    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
End Sub

Private Sub readxreturnx5(i As Integer)
' DUMMY ROUTINE for generating random pressure/flow readings

x5 = ((5 - -5 + 1) * Rnd + -5)

End Sub
