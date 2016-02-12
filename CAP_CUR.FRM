VERSION 5.00
Begin VB.Form Curve 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Polynomial Curve Fitting"
   ClientHeight    =   3615
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   8670
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3615
   ScaleWidth      =   8670
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   8415
      Begin VB.CommandButton Begin 
         Caption         =   "&BEGIN"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6720
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Help 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6720
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Cancel 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6720
         TabIndex        =   6
         Top             =   1710
         Width           =   1455
      End
      Begin VB.TextBox Num_Points 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         Top             =   870
         Width           =   585
      End
      Begin VB.TextBox Max_Distance 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5160
         TabIndex        =   2
         Top             =   1380
         Width           =   1395
      End
      Begin VB.CommandButton AddFile 
         Caption         =   "&Add File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   6720
         TabIndex        =   8
         Top             =   2370
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   420
         ScaleHeight     =   465
         ScaleWidth      =   6165
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2550
         Width           =   6225
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1980
         ScaleHeight     =   1395
         ScaleWidth      =   3915
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   510
         Width           =   3975
      End
      Begin VB.TextBox Perc_Error 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5130
         TabIndex        =   0
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Percentage error allowed (0.001-100):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   10
         Top             =   480
         Width           =   4860
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Files Selected to Average (limit 5):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         TabIndex        =   9
         Top             =   270
         Width           =   5460
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Number of points to use for fitting (3-10):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   11
         Top             =   960
         Width           =   4860
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Max. distance between any two data points in"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   4965
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Percent Complete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2460
         TabIndex        =   14
         Top             =   2310
         Width           =   2835
      End
   End
End
Attribute VB_Name = "Curve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim file_uses_fluid_sensor%
Dim AveReport As Integer
Dim Ave_File$(5)
Dim Cancelled As Integer
Dim AveFileType%
Dim ts$(41)                     ' Text strings for this form
Dim averaging As Boolean        ' Flag used in capcurve to see if routine is being called by the data averager

Private Sub AddFile_Click()

    Dim r As Integer
    
    If AveReport = 0 Then
        MsgBox ts$(7) + vbCrLf + ts$(2), 64, ts$(3)    '"All files should be of the same type."/"The first selected file will determine the type."/"Data Averager"
    End If
    AveReport = AveReport + 1
    GetaFile Me.hwnd
    file$ = fsel_return$
    If file$ = "" Then Exit Sub
    r = TestFile()
    If r = -100 Then
        MsgBox ts$(4), 0, ts$(3)            ' "Data Averager"/"Averaging cannot be done with Pressure Hold Test files."
    End If
    If r = 1 Then
        MsgBox ts$(5), 0, ts$(3)            ' "The selected file is a Capillary Flow test data file."/"Data Averager"
    End If
    If r = 2 Then
        MsgBox ts$(6), 0, ts$(3)            ' "The selected file is a Permeability data file."/"Data Averager"
    End If
    If (r <> AveFileType%) And (AveReport > 1) Then
        MsgBox ts$(7), 64, ts$(3)           ' "All files must be of the same type."/"Data Averager"
        AveReport = AveReport - 1
        Exit Sub
    End If
    If (fsel_return$ = "") Or (r = -9) Then
        AveReport = AveReport - 1
        Exit Sub
    End If
    If AveReport = 1 Then AveFileType% = r
    Ave_File$(AveReport) = fsel_return$
    Picture3.Print fsel_return$
    If AveReport = 5 Then
        AddFile.Enabled = False
        MsgBox ts$(8) + vbCrLf + "        " + ts$(9), 64, ts$(3)       ' "The maximum number of reports is selected."/"Press BEGIN to continue."/"Data Averager"
    End If
    
End Sub

Private Sub Begin_Click()

    Dim n As Integer
    Dim MaxPSI As Single
    Dim ER As Single
    
    Cancelled = False
    If Curve_Ave = 2 Then
        If AveReport < 1 Then
            MsgBox ts$(10), 0, ts$(3)                   ' "You must select more than 1 file."/"Data Averager"
            Exit Sub
        End If
        Data_Ave
        On Error Resume Next
        Unload Msgform
        On Error GoTo 0
        Exit Sub
    End If

    ER = myVal(Perc_Error.Text)
    If (ER < 0.0001) Or (ER > 100) Then
        Perc_Error.Text = ""
        Perc_Error.SetFocus
        Exit Sub
    Else
        ER = ER / 100
    End If

    n = Int(val(Num_Points.Text))
    If n < 3 Or n > 10 Then
        Num_Points.Text = ""
        Num_Points.SetFocus
        Exit Sub
    End If

    MaxPSI = myVal(Max_Distance.Text) / PCNV
    'If (MaxPsi < .0001 / PCNV) Or (MaxPsi > PY2(0) / PCNV) Then
    If (MaxPSI < 0.0001 / PCNV) Then
        MsgBox ts$(11) + " = " + str$(0.0001 / PCNV), 0, ts$(12) ' "The minimum value for pressure interval"/"Curve Fit"
        Max_Distance.Text = ""
        Max_Distance.SetFocus
        Exit Sub
    End If
    curve_perc = ER * 100
    curve_nump = n
    curve_maxd = MaxPSI
    save_user_global_stuff
    CapCurve n, ER, MaxPSI
    Me.MousePointer = 0
    
End Sub

Private Sub cancel_Click()

    Dim r As Long
    Dim numdata&
    
    numdata& = 0
    Cancelled = True
    On Error Resume Next
    r = WinHelp(hwnd, HelpFile$, Help_Quit, ByVal numdata&)
    On Error GoTo 0
    Unload Me

End Sub

Public Sub CapCurve(NX As Integer, ER As Single, MaxPSI As Single, Optional altFileName As String = "X")
' 6.71.70 rewrote to use common cft file routines - now supports temperature data
' Version 6.71.38 and higher implements this routine as an optional automatic procedure
' at the end of a test. We need to be able to call it with a file that's not selected
' explicitly by the user and that runs without user intervention. altFileName contains the file
' to use.

    ' Input variables:
    ' NX = number of points used in fitting function
    ' ER = allowable percentage error
    ' MaxPSI = maximum distance between adjacent points in PSI
    
    
    Dim DF As tcftfile ' the data file
    
    Dim i%, j%, k%                  ' loop counters
    
    Dim Q0 As Single, P0 As Single            ' Initial pressure and flow values in the file
    Dim QW As Single, PW As Single            ' Current flow and pressure values input from the file (e.g. each point in turn)
    Dim array_qw() As Single, array_pw() As Single  ' Subsets of the wet and dry flow that are used in the
                                                    ' calculations: sized according to NX
    
    ' Flow and pressure values have a three-level "stack" so we can watch
    ' the trends. Assuming an array QA of flow values with index n,
    ' QW = QA(n), LQW = QA(n-1), LlQW = QA(n-2), LLLQW = QA(n-3). Likewise
    ' for pressure.
    Dim LQW As Single, LLQW As Single, LLLQW As Single
    Dim LPW As Single, LLPW As Single, LLLPW As Single
    
    ' Calculation variables - these are used in the mathematics
    Dim P As Single, Q As Single                ' Current *calculated* flow and pressure values
    Dim XMIN As Single, XMAX As Single          ' Minimum and maximum values used in fitting subroutine
    Dim YMIN As Single, YMAX As Single          ' Minimum and maximum values used in fitting subroutine
    Dim fitp As Single                          ' Fitted pressure point
    Dim fitq As Single                          ' Fitted flow point
    
    
    Dim n As Integer
   ' Dim RI As Integer
    Dim i1 As Integer   ' some kind of array index/counter
    Dim IX As Integer   ' while loop counter
    Dim dirx As Integer ' one-time indicator
    Dim ITER As Integer ' one-time while loop counter
    Dim NITER As Integer ' one-time while loop counter
    Dim NITERM1 As Integer ' one-time while loop counter

    Begin.Enabled = False   ' Disable the "begin" button in the form
    ReDim array_qw(20), array_pw(20), xp1(100), cs(100)

' ************** Get input data file ****************
    ' Curve_Ave = 1 --> curve fitting; = 2 --> averaging
    If autoCurveFit And altFileName <> "X" Then            ' We're calling from do_final_copy
        file$ = altFileName
    ElseIf Curve_Ave <> 2 Then
        ' Not averaging; set filebox title based on type of machine
        fsel_name$ = ""
        If GasPerm Then
            fsel_title$ = ts$(13) + " " + ts$(18)       ' "Gas Permeameter Data File."
        ElseIf CFAnal Then
            fsel_title$ = ts$(14) + " " + ts$(18)       ' "CFA Data File."
        ElseIf itester Then
            fsel_title$ = ts$(15) + " " + ts$(18)       ' "Integrity Tester Data File."
        ElseIf BPTester Then
            ' 6.71.38x begin
            If hydrohead_exclusive Then
                fsel_title$ = ts$(41)                   ' "Hydrohead Tester Data File"
            Else
                fsel_title$ = ts$(16) + " " + ts$(18)       ' "Bubble Point Tester Data File"
            End If
            ' 6.71.38x end
        Else
            fsel_title$ = ts$(17) + " " + ts$(18)       ' Capillary Flow Data File."
        End If
        fsel_path$ = path(0)
        fsel_io = True
        fsel Me.hwnd
        If fsel_return$ = "" Then
            Begin.Enabled = True
            Begin.SetFocus
            Me.MousePointer = 0
            On Error GoTo 0
            Exit Sub
        End If
        path(0) = fsel_path$
        file$ = fsel_return$
    Else
        file$ = EXE_Path$ + "Tempsort.tmp"
    End If
' ************** end get input data file **************
    
    If Not autoCurveFit Then
        ' Disable settings in form
        Perc_Error.Enabled = False
        Max_Distance.Enabled = False
        Num_Points.Enabled = False
    End If
    
    ' new: read in data file using common routine
    load_cftfile DF, file$
    ' check some things
    ' note - fluid sensor is not supported in load_cftfile - we haven't used it in a long long time
    
    If DF.cfttype = ect_phold Then
        'MsgBox ts$(20), 0, ts$(11)         ' "Curve Fitting cannot be done on Pressure Hold Test Data Files."/"Curve Fit"
        Begin.Enabled = True
        Begin.SetFocus
        Me.MousePointer = 0
        On Error GoTo 0
        Exit Sub
    End If
    If DF.cfttype = ect_bubblepoint Then
        MsgBox ts$(21), 0, ts$(11)          ' "Curve fitting cannot be run on a Bubble Point data file!"/"Curve Fit"
        Begin.Enabled = True
        Begin.SetFocus
        Me.MousePointer = 0
        On Error GoTo 0
        Exit Sub
    End If
    ' Check for user intervention
    If Cancelled Then
        Me.MousePointer = 0
        On Error GoTo 0
        Exit Sub
    End If
        
        ' Start processing the actual data
    Q0 = DF.Q0
    P0 = DF.P0
    ' initial scan to throw out some data points
    i% = 1
    While i% <= DF.num_wet_points
        QW = DF.wet_data.a(i%).fl
        PW = DF.wet_data.a(i%).pr
        ' throw out current data point if:
        '  it is the first point and pressure or flow are negative
        '  it is not the first point and we are curve fitting and the pressure is falling back
        '  it is not the first or second point and the flow hasn't changed for last three points
        If (i% = 1 And (PW < P0 Or QW < Q0)) Or _
           (i% > 1 And Curve_Ave <> 2 And PW < LPW) Or _
           (i% > 2 And QW = LQW And QW = LLQW) Then
            ' first, handle special case of throwing out middle point
            If (i% > 2 And QW = LQW And QW = LLQW) Then
                LQW = QW
                LPW = PW
                DF.wet_data.a(i% - 1) = DF.wet_data.a(i%)
                If DF.temperaturesensor = 1 Then
                    DF.wet_ts1_data.a(i% - 1) = DF.wet_ts1_data.a(i%)
                ElseIf DF.temperaturesensor = 2 Then
                    DF.wet_ts2_data.a(i% - 1) = DF.wet_ts2_data.a(i%)
                End If
            End If
            ' next, remove data point i% from data and shove everything else down by one
            For j% = i% + 1 To DF.num_wet_points
                DF.wet_data.a(j% - 1) = DF.wet_data.a(j%)
                If DF.temperaturesensor = 1 Then
                    DF.wet_ts1_data.a(j% - 1) = DF.wet_ts1_data.a(j%)
                ElseIf DF.temperaturesensor = 2 Then
                    DF.wet_ts2_data.a(j% - 1) = DF.wet_ts2_data.a(j%)
                End If
            Next j%
            DF.num_wet_points = DF.num_wet_points - 1
        Else
            ' Shift flow and pressure points back: P and Q have a three-level "stack" of previous values
            LPW = PW
            LLQW = LQW: LQW = QW
            i% = i% + 1
        End If
    Wend
    ' next scan to add in some more points, if any are too close together
    i% = 1
    While i% <= DF.num_wet_points
        QW = DF.wet_data.a(i%).fl
        PW = DF.wet_data.a(i%).pr
        If i% > 1 And Abs(LPW - PW) >= MaxPSI Then
            ' add in at least one more point
            n = Int(Abs(LPW - PW) / MaxPSI)     ' Find the number of interpolated points
            ' make sure there is room for it
            pf_size_check DF.wet_data, DF.num_wet_points + n
            If DF.temperaturesensor = 1 Then
                ts1_size_check DF.wet_ts1_data, DF.num_wet_points + n
            ElseIf DF.temperaturesensor = 2 Then
                ts2_size_check DF.wet_ts2_data, DF.num_wet_points + n
            End If
            ' shove everything else up by n
            For k% = DF.num_wet_points To i% Step -1
                DF.wet_data.a(k% + n) = DF.wet_data.a(k%)
                If DF.temperaturesensor = 1 Then
                    DF.wet_ts1_data.a(k% + n) = DF.wet_ts1_data.a(k%)
                ElseIf DF.temperaturesensor = 2 Then
                    DF.wet_ts2_data.a(k% + n) = DF.wet_ts2_data.a(k%)
                End If
            Next k%
            DF.num_wet_points = DF.num_wet_points + n
            For j% = 1 To n
                ' Calculate fractional flow and pressure values
                DF.wet_data.a(i% + j% - 1).pr = LPW + (PW - LPW) * j / (n + 1)
                DF.wet_data.a(i% + j% - 1).fl = LQW + (QW - LQW) * j / (n + 1)
                ' just in case there is any time data, this can't hurt
                DF.wet_data.a(i% + j% - 1).time = DF.wet_data.a(i% - 1).time + _
                  (DF.wet_data.a(i% + n).time - DF.wet_data.a(i% - 1).time) * j / (n + 1)
                If DF.temperaturesensor = 1 Then
                    DF.wet_ts1_data.a(i% + j% - 1) = DF.wet_ts1_data.a(i% - 1) + _
                      (DF.wet_ts1_data.a(i% + n) - DF.wet_ts1_data.a(i% - 1)) * j / (n + 1)
                ElseIf DF.temperaturesensor = 2 Then
                    DF.wet_ts2_data.a(i% + j% - 1).temp1 = DF.wet_ts2_data.a(i% - 1).temp1 + _
                        (DF.wet_ts2_data.a(i% + n).temp1 - DF.wet_ts2_data.a(i% - 1).temp1) * j / (n + 1)
                    DF.wet_ts2_data.a(i% + j% - 1).temp2 = DF.wet_ts2_data.a(i% - 1).temp2 + _
                        (DF.wet_ts2_data.a(i% + n).temp2 - DF.wet_ts2_data.a(i% - 1).temp2) * j / (n + 1)
                End If
            Next j      ' Next fractional point
            i% = i% + n
        End If
        ' this point is good (or is the first, which has to be good)
        LQW = QW
        LPW = PW
        i% = i% + 1
        ' Be nice to the system and check for user intervention
        DoEvents
        If Cancelled Then
            Me.MousePointer = 0
            On Error GoTo 0
            Exit Sub
        End If
    Wend
                
    ' throw out any trailing negative values
    If DF.num_wet_points > 0 Then
        While (DF.wet_data.a(DF.num_wet_points).pr < P0 Or _
                DF.wet_data.a(DF.num_wet_points).fl < Q0) And _
              DF.num_wet_points > 0
            DF.num_wet_points = DF.num_wet_points - 1
        Wend
    End If
              
    ' redo all of above for the dry section of the data file
    ' initial scan to throw out some data points
    i% = 1
    While i% <= DF.num_dry_points
        QW = DF.dry_data.a(i%).fl
        PW = DF.dry_data.a(i%).pr
        ' throw out current data point if:
        '  it is the first point and pressure or flow are negative
        '  it is not the first point and we are curve fitting and the pressure is falling back
        '  it is not the first or second point and the flow hasn't changed for last three points
        If (i% = 1 And (PW < P0 Or QW < Q0)) Or _
           (i% > 1 And Curve_Ave <> 2 And PW < LPW) Or _
           (i% > 2 And QW = LQW And QW = LLQW) Then
            ' first, handle special case of throwing out middle point
            If (i% > 2 And QW = LQW And QW = LLQW) Then
                LQW = QW
                LPW = PW
                DF.dry_data.a(i% - 1) = DF.dry_data.a(i%)
                If DF.temperaturesensor = 1 Then
                    DF.ts1_data.a(i% - 1) = DF.ts1_data.a(i%)
                ElseIf DF.temperaturesensor = 2 Then
                    DF.ts2_data.a(i% - 1) = DF.ts2_data.a(i%)
                End If
            End If
            ' next, remove data point i% from data and shove everything else down by one
            For j% = i% + 1 To DF.num_dry_points
                DF.dry_data.a(j% - 1) = DF.dry_data.a(j%)
                If DF.temperaturesensor = 1 Then
                    DF.ts1_data.a(j% - 1) = DF.ts1_data.a(j%)
                ElseIf DF.temperaturesensor = 2 Then
                    DF.ts2_data.a(j% - 1) = DF.ts2_data.a(j%)
                End If
            Next j%
            DF.num_dry_points = DF.num_dry_points - 1
        Else
            ' Shift flow and pressure points back: P and Q have a three-level "stack" of previous values
            LPW = PW
            LLQW = LQW: LQW = QW
            i% = i% + 1
        End If
    Wend
    ' next scan to add in some more points, if any are too close together
    i% = 1
    While i% <= DF.num_dry_points
        QW = DF.dry_data.a(i%).fl
        PW = DF.dry_data.a(i%).pr
        If i% > 1 And Abs(LPW - PW) >= MaxPSI Then
            ' add in at least one more point
            n = Int(Abs(LPW - PW) / MaxPSI)     ' Find the number of interpolated points
            ' make sure there is room for it
            pf_size_check DF.dry_data, DF.num_dry_points + n
            If DF.temperaturesensor = 1 Then
                ts1_size_check DF.ts1_data, DF.num_dry_points + n
            ElseIf DF.temperaturesensor = 2 Then
                ts2_size_check DF.ts2_data, DF.num_dry_points + n
            End If
            ' shove everything else up by n
            For k% = DF.num_dry_points To i% Step -1
                DF.dry_data.a(k% + n) = DF.dry_data.a(k%)
                If DF.temperaturesensor = 1 Then
                    DF.ts1_data.a(k% + n) = DF.ts1_data.a(k%)
                ElseIf DF.temperaturesensor = 2 Then
                    DF.ts2_data.a(k% + n) = DF.ts2_data.a(k%)
                End If
            Next k%
            DF.num_dry_points = DF.num_dry_points + n
            For j% = 1 To n
                ' Calculate fractional flow and pressure values
                DF.dry_data.a(i% + j% - 1).pr = LPW + (PW - LPW) * j / (n + 1)
                DF.dry_data.a(i% + j% - 1).fl = LQW + (QW - LQW) * j / (n + 1)
                ' just in case there is any time data, this can't hurt
                DF.dry_data.a(i% + j% - 1).time = DF.dry_data.a(i% - 1).time + _
                  (DF.dry_data.a(i% + n).time - DF.dry_data.a(i% - 1).time) * j / (n + 1)
                If DF.temperaturesensor = 1 Then
                    DF.ts1_data.a(i% + j% - 1) = DF.ts1_data.a(i% - 1) + _
                      (DF.ts1_data.a(i% + n) - DF.ts1_data.a(i% - 1)) * j / (n + 1)
                ElseIf DF.temperaturesensor = 2 Then
                    DF.ts2_data.a(i% + j% - 1).temp1 = DF.ts2_data.a(i% - 1).temp1 + _
                        (DF.ts2_data.a(i% + n).temp1 - DF.ts2_data.a(i% - 1).temp1) * j / (n + 1)
                    DF.ts2_data.a(i% + j% - 1).temp2 = DF.ts2_data.a(i% - 1).temp2 + _
                        (DF.ts2_data.a(i% + n).temp2 - DF.ts2_data.a(i% - 1).temp2) * j / (n + 1)
                End If
            Next j      ' Next fractional point
            i% = i% + n
        End If
        ' this point is good (or is the first, which has to be good)
        LQW = QW
        LPW = PW
        i% = i% + 1
        ' Be nice to the system and check for user intervention
        DoEvents
        If Cancelled Then
            Me.MousePointer = 0
            On Error GoTo 0
            Exit Sub
        End If
    Wend
                
    ' throw out any trailing negative values
    If DF.num_dry_points > 0 Then
        While (DF.dry_data.a(DF.num_dry_points).pr < P0 Or _
                DF.dry_data.a(DF.num_dry_points).fl < Q0) And _
              DF.num_dry_points > 0
            DF.num_dry_points = DF.num_dry_points - 1
        Wend
    End If
    
    ' if we have both wet and dry, throw out any initial wet points that are below the initial dry point
    If DF.num_dry_points > 0 And DF.num_wet_points > 0 Then
        i% = 1
        Do
            If DF.wet_data.a(i%).pr > DF.dry_data.a(1).pr Then Exit Do
            i% = i% + 1
        Loop Until i% > DF.num_wet_points
        If i% > DF.num_wet_points Then
            DF.num_wet_points = 0
        ElseIf i% > 1 Then
            j% = 1
            For k% = i% To DF.num_wet_points
                DF.wet_data.a(j%) = DF.wet_data.a(k%)
                If DF.temperaturesensor = 1 Then
                    DF.wet_ts1_data.a(j%) = DF.wet_ts1_data.a(k%)
                ElseIf DF.temperaturesensor = 2 Then
                    DF.wet_ts2_data.a(j%) = DF.wet_ts2_data.a(k%)
                End If
                j% = j% + 1
            Next k%
            DF.num_wet_points = DF.num_wet_points - i% + 1
        End If
    End If
    DoEvents
    If Cancelled Then
        Me.MousePointer = 0
        On Error GoTo 0
        Exit Sub
    End If
'***************** end of data input section! ********************************

' **************** Data validation and check with user before starting *******
4000  ' RI = 1
    ' Check data file to make sure we can do the fit
    If Curve_Ave = 2 Then GoTo NO_AVE   ' i.e., if we're averaging, not fitting
    If (DF.num_wet_points > 0 And DF.num_wet_points < 3) Or _
       (DF.num_dry_points > 0 And DF.num_dry_points < 2) Then ' Make sure we have enough points. If not, get out!
        MsgBox ts$(22) + " : " + file$ + vbCrLf + ts$(23), 0, ts$(24)     ' "Cannot Curve Fit File"/"Not Enough Points in file."/"Data Error"
        Begin.Enabled = True
        Begin.SetFocus
        Perc_Error.Enabled = True
        Max_Distance.Enabled = True
        Num_Points.Enabled = True
        Me.MousePointer = 0
        On Error GoTo 0
        Exit Sub
    End If

    If Not autoCurveFit Then        ' No display to user if we're calling from do_final_copy
        If DF.num_wet_points > 0 And DF.num_dry_points > 0 Then
            ' Tell the user how much data there is
            reply% = MsgBox(str$(DF.num_wet_points) + " " + ts$(25) + vbCrLf + str$(DF.num_dry_points) + " " + ts$(26) + vbCrLf + ts$(27) + " : " + file$, 1, ts$(28))       ' "wet data points and"/"dry data points in"/"File"/"Data"
        ElseIf DF.num_dry_points > 0 Then
            reply% = MsgBox(str$(DF.num_dry_points) + " " + ts$(29) + " : " + file$, 1, ts$(28))      '"data points in File"/"Data"
        Else
            reply% = MsgBox(str$(DF.num_wet_points) + " " + ts$(29) + " : " + file$, 1, ts$(28))      '"data points in File"/"Data"
        End If
        ' Check to see if the user cancels out
        If reply% = 2 Then
            Begin.Enabled = True
            Begin.SetFocus
            Perc_Error.Enabled = True
            Max_Distance.Enabled = True
            Num_Points.Enabled = True
            Me.MousePointer = 0
            On Error GoTo 0
            Exit Sub
        End If
        
    End If
' *************** Done with user validation and data check **************************

' *************** Begin data processing section *************************************
NO_AVE:
    ' Make sure we have more data points than the selected number of points used in fitting (NX).
    ' If not, reduce NX to the minimum number of points
    If DF.num_wet_points > 0 And DF.num_wet_points < NX Then
        NX = DF.num_wet_points
    End If
    If DF.num_dry_points > 0 And DF.num_dry_points < NX Then
        NX = DF.num_dry_points
    End If

    If Cancelled Then
        Me.MousePointer = 0
        On Error GoTo 0
        Exit Sub
    End If
    DoEvents
    
    If DF.num_wet_points > 0 Then
        ' Load the set of points to use in fitting
        For i% = 1 To NX
            array_pw(i%) = DF.wet_data.a(i%).pr
            array_qw(i%) = DF.wet_data.a(i%).fl
        Next i%
    
        i1 = 0: IX = 2: dirx = 1
        ' Update the display if appropriate
        If Not autoCurveFit Then
            Picture1.Cls
            If DF.num_dry_points > 0 Then
                Label1.Caption = UCase$(ts$(30)) + " - " + ts$(32)       ' "WET"/"Percent Complete'
            Else
                Label1.Caption = ts$(32)                                 ' "Percent Complete"
            End If
            Me.Refresh
            Picture1.Scale (0, 0)-(DF.num_wet_points, 100)
        End If
    
        ' Process the first data set!
        While IX < DF.num_wet_points
            If Not autoCurveFit Then Picture1.Line (0, 0)-(IX, 100), &HFF0000, BF
            GoSub 8000                  ' <----- This is the call to the real curve fitting routine
            DF.wet_data.a(IX).pr = fitp
            DF.wet_data.a(IX).fl = fitq
            IX = IX + 1
            If (1 + NX) / 2 < IX - i1 And NX + i1 < DF.num_wet_points Then
                For i% = 1 To NX - 1
                    array_pw(i%) = array_pw(i% + 1): array_qw(i%) = array_qw(i% + 1)
                Next i%
                i1 = i1 + 1
                array_pw(NX) = DF.wet_data.a(NX + i1).pr
                array_qw(NX) = DF.wet_data.a(NX + i1).fl
            End If
        Wend
        If Not autoCurveFit Then Picture1.Line (0, 0)-(DF.num_wet_points, 100), &HFF0000, BF
    End If
    ' Check for user and system input
    DoEvents
    If Cancelled Then
        Me.MousePointer = 0
        On Error GoTo 0
        Exit Sub
    End If
    

    If DF.num_dry_points > 0 Then

        ' Load the points used for fitting into the arrays
        For i% = 1 To NX
            array_pw(i%) = DF.dry_data.a(i%).pr
            array_qw(i%) = DF.dry_data.a(i%).fl
        Next i%
        
        ' start at second data point (IX=2)
        ' with an offset of 0 (i1=0)
        ' and data is going in normal direction (dirx=1)
        ' Note that previous version used dirx=0 because dry curve was running backwards
        ' Current version already has dry curve reversed from the data file, so it is running forwards
        IX = 2: i1 = 0: dirx = 1
        ' Refresh the display if appropriate
        If Not autoCurveFit Then
            Picture1.Cls
            If DF.num_wet_points > 0 Then
                Label1.Caption = UCase$(ts$(31)) + " - " + ts$(32)      ' "DRY - Percent Complete"
            Else
                Label1.Caption = ts$(32)                                 ' "Percent Complete"
            End If
            Me.Refresh
            Picture1.Scale (0, 0)-(DF.num_dry_points, 100)
        End If
    
        ' Process the second data set
        While IX < DF.num_dry_points
            If Not autoCurveFit Then Picture1.Line (0, 0)-(IX, 100), &HFF0000, BF
            GoSub 8000                  ' <----- This is the call to the real curve fitting routine
            
            DoEvents
            If Cancelled Then
                Me.MousePointer = 0
                On Error GoTo 0
                Exit Sub
            End If
        
            DF.dry_data.a(IX).pr = fitp
            DF.dry_data.a(IX).fl = fitq
        
            IX = IX + 1
            If (1 + NX) / 2 < IX - i1 And NX + i1 < DF.num_dry_points Then
                For i% = 1 To NX - 1
                    array_pw(i%) = array_pw(i% + 1)
                    array_qw(i%) = array_qw(i% + 1)
                Next i%
                i1 = i1 + 1
                array_pw(NX) = DF.dry_data.a(NX + i1).pr
                array_qw(NX) = DF.dry_data.a(NX + i1).fl
            End If
        Wend
    
        If Not autoCurveFit Then Picture1.Line (0, 0)-(DF.num_dry_points, 100), &HFF0000, BF
    End If

' ********************* done with fitting! ************************

    ' Get destination filename from user if appropriate
    If Not averaging And Not autoCurveFit Then      ' averaging is true if function is being called by
                                                    ' the data averager routine as an intermediate step
                                                    ' which the user should not see; likewise, autocurvefit
                                                    ' is true if we are calling this at the end of a test
        fsel_name$ = ""
        If Curve_Ave <> 2 Then
            fsel_title$ = ts$(33)   ' "Curve-Fit Data File SAVE AS"
        Else
            fsel_title$ = ts$(34)   ' "Averaged Data File SAVE AS"
        End If
        
        fsel_path$ = path(0)
        fsel_io = False
        fsel Me.hwnd
        
        If fsel_return$ = "" Then
            Begin.Enabled = True
            Begin.SetFocus
            'Unload me
            Me.MousePointer = 0
            On Error GoTo 0
            Exit Sub
        End If
        
        path(0) = fsel_path$
        file$ = fsel_return$
    ElseIf averaging Then
        file$ = EXE_Path + "avetemp.tmp"    ' Write to a temporary file to be retrieved by the averaging routine
    Else
        file$ = altFileName             ' Write back to the same file (which should be tempdata2.cft)
    End If
    save_cftfile DF, file$

    ' Reset the window (if appropriate)
    If Not autoCurveFit Then
        Begin.Enabled = True
        Begin.SetFocus
        Perc_Error.Enabled = True
        Max_Distance.Enabled = True
        Num_Points.Enabled = True
        Me.MousePointer = 0
    End If
    
    'Unload me
    On Error GoTo 0
    Exit Sub    ' Done!

'************ end of main routine! ******************************************
'****************************************************************************

'
'
' *************** Beginning of the real curve fit routine *******************
'
'
8000 ' given points array_pw(1..n) and array_qw(1..n), find point on best fit curve
    ' that corresponds to array_pw(ix-i1) and set fitp, fitw to it. (The "n" in this
    ' case is NX, the number of points selected for fitting by the user.)

    ' Set initial bounds to the first flow and pressure values
    XMIN = array_pw(1)
    YMIN = array_qw(1)
    XMAX = XMIN
    YMAX = YMIN
    
    ' Find the minimum and maximum flow and pressure values in the arrays. (X is pressure,
    ' Y is flow)
    For j = 2 To NX         ' For every point in the arrays higher than 1 ...
        
        ' Be nice
        DoEvents
        If Cancelled Then
            Me.MousePointer = 0
            On Error GoTo 0
            Exit Sub
        End If
        
        If array_pw(j) < XMIN Then XMIN = array_pw(j)
        If array_pw(j) > XMAX Then XMAX = array_pw(j)
        If array_qw(j) < YMIN Then YMIN = array_qw(j)
        If array_qw(j) > YMAX Then YMAX = array_qw(j)
        
    Next j
    
    ' Make sure we have some minimum interval
    If YMIN = YMAX Then YMAX = YMAX + 0.000001
    If XMIN = XMAX Then XMAX = XMAX + 0.000001
    
    ' Load the arrays into the xp1 and cs arrays, shifting the indices by one
    For j = 1 To NX
        xp1(j - 1) = array_pw(j): cs(j - 1) = array_qw(j)
    Next j

8020 ' now find "center" point

    ' NITER = maximum iterations; ITER = current iteration
    ITER = 1: NITER = NX: NITERM1 = 0
    
    ' Add values to the end of the xp and cs arrays (same number of points
    ' as were initially put in?) Center value is left empty.
    While ITER < NX
    
        For j = 0 To NX - 1 - ITER
            xp1(NITER + j) = (xp1(NITERM1 + j) + xp1(NITERM1 + j + 1)) / 2
            cs(NITER + j) = (cs(NITERM1 + j) + cs(NITERM1 + j + 1)) / 2
        Next j
        
        ITER = ITER + 1: NITERM1 = NITER: NITER = NITER + NX
    
    Wend
    
    If Abs((array_pw(IX - i1) - xp1(NITERM1)) / (XMAX - XMIN)) < ER Then
        fitq = cs(NITERM1): fitp = xp1(NITERM1)
        Return
    End If
    
    ' here we need to force a fit if you can't get any closer
    If xp1(1) = xp1(NX) Then
        fitq = cs(NITERM1): fitp = xp1(NITERM1)
        Return
    End If
    
    DoEvents
    If Cancelled Then
        Me.MousePointer = 0
        On Error GoTo 0
        Exit Sub
    End If
    
    If dirx = 1 Then If array_pw(IX - i1) < xp1(NITERM1) Then GoTo 8100 Else GoTo 8200
    If array_pw(IX - i1) < xp1(NITERM1) Then GoTo 8200 Else GoTo 8100

8100 ' use first half
    For j = 1 To NX - 1
        xp1(j) = xp1(j * NX)
        cs(j) = cs(j * NX)
    Next j
    GoTo 8020

8200 ' use second half
    For j = 0 To NX - 2
        xp1(j) = xp1(NX * (NX - j - 1) + j)
        cs(j) = cs(NX * (NX - j - 1) + j)
    Next j
    GoTo 8020

Curve_Error:
    MsgBox ts$(35), 0, ts$(36)      '"There is a problem with your data file."/"File Error"
    Begin.Enabled = True
    Begin.SetFocus
    Perc_Error.Enabled = True
    Max_Distance.Enabled = True
    Num_Points.Enabled = True
    Resume C_Finis
    
C_Finis:
    On Error GoTo 0
    
End Sub

Private Sub Data_Ave()
' Create an average wet curve and dry curve. Do this by:
'   1)  Averaging the selected dry curves together
'   2)  Calculating the cumulative filter flow curve for each data file
'   3)  Using these to create an average FF curve
'   4)  Using the averaged dry curve and the averaged FF curve to calculate a corresponding
'       wet curve.
'   5)  Saving the resulting wet curve and dry curve into a new data file

    Dim datafiles() As tcftfile      ' Array to hold pressure/flow information
    Dim CFF() As pf_array            ' Array to hold cumulative filter flow for each file
    Dim CFFallflows As pf_array      ' Array to hold CFF for all files
    Dim dryflows As pf_array         ' Array to hold combined dry flows
    Dim i As Integer, j As Integer, cffcount As Integer, drycount As Integer
    Dim QD As Single, QW As Single   ' Current dry and wet flow values
    Dim LQD As Single                ' Previous dry flow value
    Dim P As Single                  ' Current wet pressure values
    Dim LPD As Single                ' Previous dry pressure value
    Dim DI As Integer
    Dim tempfile As tcftfile         ' temporary CFT file for curve fitting
    Dim finalfile As tcftfile        ' Final result of this whole mess, with the averaged dry flow
                                     ' values and the calculated wet flow values
                                     

    ReDim datafiles(AveReport)                      ' AveReport = # of files selected for averaging
    ReDim CFF(AveReport)
    
    ' Read in data files
    For i = 1 To AveReport
        load_cftfile datafiles(i), Ave_File$(i)     ' Load each data file into a tcftfile
    Next i
    
    ' Calculate CFF for each and store in a pf_array structure (this section was copied
    ' almost verbatim from PrintReport in CapRep, apart from variable names)
    For i = 1 To AveReport
        DI = 2
        CFF(i).s = datafiles(i).num_wet_points          ' first point doesn't count
        ReDim CFF(i).a(CFF(i).s)
        For j = 1 To CFF(i).s                          ' Loop through total number of wet points
            If datafiles(i).num_dry_points > 0 Then
                ' While dry pressure is less than wet pressure ...
                While datafiles(i).dry_data.a(DI).pr < datafiles(i).wet_data.a(j).pr And DI < datafiles(i).num_dry_points
                    LPD = datafiles(i).dry_data.a(DI).pr        ' previous dry pressure
                    LQD = datafiles(i).dry_data.a(DI).fl        ' and flow data point
                    DI = DI + 1
                Wend
            End If
            P = datafiles(i).wet_data.a(j).pr                   ' current wet pressure
            QW = datafiles(i).wet_data.a(j).fl                  ' current wet flow
            If datafiles(i).num_dry_points > 0 Then
                If datafiles(i).dry_data.a(DI).pr = P Then      ' if dry and wet pressures are equal...
                    QD = datafiles(i).dry_data.a(DI).fl
                Else
                    ' Interpolate dry data to find a value for wet pressure
                    QD = LQD + (P - LPD) / (datafiles(i).dry_data.a(DI).pr - LPD) * (datafiles(i).dry_data.a(DI).fl - LQD)
                End If
                CFF(i).a(j).fl = QW / QD * 100                  ' CFF from wet and dry flows
                CFF(i).a(j).pr = P
                If CFF(i).a(j).fl > 100 Then CFF(i).a(j).fl = 100   ' max out at 100%, naturally
            End If
        Next j
    Next i
    
    ' Average CFF and dry curves (independently)
    dryflows.s = 0
    CFFallflows.s = 0
    For i = 1 To AveReport
        dryflows.s = dryflows.s + datafiles(i).num_dry_points
        CFFallflows.s = CFFallflows.s + CFF(i).s
    Next i
    ReDim CFFallflows.a(CFFallflows.s)
    ReDim dryflows.a(dryflows.s)
    cffcount = 0:  drycount = 0
    For i = 1 To AveReport
        For j = 1 To CFF(i).s
            cffcount = cffcount + 1
            CFFallflows.a(cffcount) = CFF(i).a(j)
        Next j
        For j = 1 To datafiles(i).num_dry_points
            drycount = drycount + 1
            dryflows.a(drycount).fl = datafiles(i).dry_data.a(j).fl
            ' when copying over the dry pressures, correct for any atmospheric pressure change
            ' between the first file and all subsequent files
            dryflows.a(drycount).pr = datafiles(i).dry_data.a(j).pr - datafiles(i).P0 + datafiles(1).P0
        Next j
    Next i
 
    QuickSort 1, CFFallflows.s, CFFallflows
    ' CFFallflows now contains the sorted CFF values
    QuickSort 1, dryflows.s, dryflows
    ' dryflows now contains the sorted dry flow values
    
    ' To use the curve fitter, we must save these values as a temporary CFT file
    ' and feed it to CapCurve
  
    ' First do some manipulating to turn the pf_array into a tcftfile
    initialize_cftfile tempfile
    tempfile = datafiles(1)                 ' Set header information to be the same
    tempfile.num_wet_points = CFFallflows.s
    tempfile.num_dry_points = dryflows.s
    ReDim tempfile.wet_data.a(CFFallflows.s)
    ReDim tempfile.dry_data.a(dryflows.s)
    For i = 1 To CFFallflows.s
        tempfile.wet_data.a(i).fl = CFFallflows.a(i).fl
        tempfile.wet_data.a(i).pr = CFFallflows.a(i).pr
    Next i
    For i = 1 To dryflows.s
        tempfile.dry_data.a(i).fl = dryflows.a(i).fl
        tempfile.dry_data.a(i).pr = dryflows.a(i).pr
    Next i
    save_cftfile tempfile, EXE_Path + "\Tempsort.tmp"
    averaging = True
    CapCurve 5, 0.1, 1
    Kill EXE_Path + "\Tempsort.tmp"
    
    ' if this is a dry curve only file, we are done
    If CFFallflows.s = 0 Then
        ' just load the temp file into the final structure in preparation for saving it
        ' to wherever the user wants it
        load_cftfile finalfile, EXE_Path + "\avetemp.tmp"
    Else
        ' otherwise, we still have some manipulating to do with the curve fit file before
        ' it is ready to save:
        ' Now we have a curve-fitted data set of the average CFF values, stored as the wet
        ' flow data in a temporary file, and a similar set of average dry flows stored as the dry
        ' data. We will load these back into CFFallflows and dryflows.
        load_cftfile tempfile, EXE_Path + "\avetemp.tmp"
        ReDim CFFallflows.a(tempfile.num_wet_points)
        ReDim dryflows.a(tempfile.num_dry_points)
        CFFallflows.s = tempfile.num_wet_points
        dryflows.s = tempfile.num_dry_points
        For i = 1 To CFFallflows.s
            CFFallflows.a(i).fl = tempfile.wet_data.a(i).fl
            CFFallflows.a(i).pr = tempfile.wet_data.a(i).pr
        Next i
        For i = 1 To dryflows.s
            dryflows.a(i).fl = tempfile.dry_data.a(i).fl
            dryflows.a(i).pr = tempfile.dry_data.a(i).pr
        Next i
        
        averaging = False
        ' Now we finally have the averaged values!
        
        ' We are almost done -- we must reverse our first procedure and back-calculate wet
        ' flow data to match the dry and CFF data. We will calculate for each point of the
        ' CFF curve and preserve only the corresponding dry curve points
        initialize_cftfile finalfile
        finalfile = datafiles(1)                    ' Use first file header info
        
        
        ReDim finalfile.wet_data.a(CFFallflows.s)
        ReDim finalfile.dry_data.a(CFFallflows.s)
        
        finalfile.wet_data.s = CFFallflows.s
        finalfile.dry_data.s = CFFallflows.s
        finalfile.num_wet_points = CFFallflows.s
        finalfile.num_dry_points = CFFallflows.s
        DI = 2
        For i = 1 To CFFallflows.s          ' Loop through all CFF values
            ' While dry pressure is less than wet pressure ...
            While (dryflows.a(DI).pr < CFFallflows.a(i).pr And DI < dryflows.s)
                LPD = dryflows.a(DI).pr
                LQD = dryflows.a(DI).fl
                DI = DI + 1
            Wend
            P = CFFallflows.a(i).pr
            QW = CFFallflows.a(i).fl    ' QW is calculated for the sake of consistency with the initial calculations, but it is actually CFF
            If dryflows.s > 0 Then
                If dryflows.a(DI).pr = P Then
                    QD = dryflows.a(DI).fl
                Else
                    QD = LQD + (P - LPD) / (dryflows.a(DI).pr - LPD) * (dryflows.a(DI).fl - LQD)
                End If
                finalfile.wet_data.a(i).fl = QW * QD / 100
                finalfile.wet_data.a(i).pr = CFFallflows.a(i).pr
                finalfile.dry_data.a(i).fl = QD
                finalfile.dry_data.a(i).pr = CFFallflows.a(i).pr
            End If
        Next i
    End If
    
    ' That's it! We should now have the appropriate values in finalfile. Just make a few tweaks
    ' to the file info ... and then save
    fsel_name$ = ""
    fsel_title$ = ts$(34)               ' "Averaged Data File SAVE AS"
    fsel_path$ = path(0)
    fsel_io = False
    fsel Me.hwnd
    If fsel_return$ = "" Then
        Exit Sub
    End If
    path(0) = fsel_path$
    file$ = fsel_return$
    save_cftfile finalfile, file$       ' EXE_Path + "\finalfile.cft"
    
End Sub

Private Sub Form_Load()

    LoadTextStrings

    TitleScrn.modifymenu(2).Enabled = False
    TitleScrn.modifymenu(4).Enabled = False
    extend_num% = 0
    If Curve_Ave = 2 Then
        Caption = ts$(38)           ' "Data Averager"
        Label2.Visible = True
        Label9.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        Perc_Error.Visible = False
        Num_Points.Visible = False
        Max_Distance.Visible = False
        AddFile.Visible = True
        Picture3.Visible = True
        AddFile.Enabled = True
        AveReport = 0
    Else
        Caption = ts$(39)           ' "Polynomial Curve Fitting"
        Label2.Visible = False
        Label9.Visible = True
        Label11.Visible = True
        Label12.Visible = True
        Perc_Error.Visible = True
        Num_Points.Visible = True
        Max_Distance.Visible = True
        AddFile.Visible = False
        Picture3.Visible = False
        Perc_Error.Text = Format$(curve_perc)
        Num_Points.Text = Format$(curve_nump)
        Max_Distance.Text = Format$(curve_maxd * PCNV)
    End If
    cancel.Tag = ""
    Label9.Caption = Label9.Caption + " " + UCase$(PU$)
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    AveReport = 0
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(1) + str$(unitnumber)  ' "Unit"
    End If
    
    averaging = False
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
End Sub

Private Sub Form_Unload(cancel As Integer)
    Me.MousePointer = 0
    TitleScrn.modifymenu(2).Enabled = True
    TitleScrn.modifymenu(4).Enabled = True
End Sub

Private Sub Help_Click()

    Dim numdata&
    Dim r As Long
    
    If Curve_Ave = 2 Then numdata& = 40 Else numdata& = 30
    r = WinHelp(hwnd, HelpFile$, Help_Context, ByVal numdata&)

End Sub
    
Private Sub Max_Distance_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 8, Asc(Decimal_Point$), 48 To 57
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub Num_Points_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub Num_Points_LostFocus()
    Num_Points.Text = Format$(val(Num_Points.Text), "#0")
End Sub

Private Sub Perc_Error_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 8, Asc(Decimal_Point$), 48 To 57
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub Perc_Error_LostFocus()
    Perc_Error.Text = Format$(val(Perc_Error.Text), "##0.000")
End Sub

Private Function TestFile()

    Dim DT$, i%, using_extra_fluid%, ID$, temp$, Filetype%
    Dim dummy As Single
    Dim Q As Single
    Dim P As Single
    Dim thickness As Single
    Dim fn1 As Integer
    Dim visc As Single
    
    fn1 = FreeFile
    Open file$ For Input As #fn1
        
    Line Input #fn1, DT$
    If DT$ = "EXTENDED" Then
        Input #fn1, i%
        While i% > 0
            Line Input #fn1, DT$
            i% = i% - 1
        Wend
        Line Input #fn1, DT$
    End If
    If DT$ = "FLUIDSENSOR" Then
        Line Input #fn1, DT$      ' fsunit$
        Line Input #fn1, DT$      ' real date
        file_uses_fluid_sensor% = True
        MsgBox ts$(19)            ' "Warning:  Fluid Sensor data will not be saved to new file"
    Else
        file_uses_fluid_sensor% = False
    End If
    If DT$ = "HOLD" Then
        TestFile = -100
        Close #fn1
        Exit Function
    End If
    Line Input #fn1, Line1$(current_unit%)
    Line Input #fn1, Line2$(current_unit%)
    Line Input #fn1, fluid$(current_unit%)
    Input #fn1, surfTen(current_unit%)
    If Sgn(surfTen(current_unit%)) = -1 Then
        Input #fn1, Diam(current_unit%)
        If Diam(current_unit%) = 0 Then
            Input #fn1, Diam(current_unit%), cyl_len(current_unit%)
        Else
            cyl_len(current_unit%) = 0
        End If
        Input #fn1, thickness
    End If
    If surfTen(current_unit%) = 0 And fluid$(current_unit%) = "X" Then
        using_extra_fluid% = True
        Line Input #fn1, fluid$(current_unit%)
        Input #fn1, visc
        fluid$(current_unit%) = Left$(str$(visc) + String$(20, " "), 20) + fluid$(current_unit%)
    Else
        using_extra_fluid% = False
    End If
    Line Input #fn1, ID$
    Line Input #fn1, temp$
    If InStr(temp$, ",") <> 0 Then
        Filetype% = 1
        Input #fn1, Q, P
        If file_uses_fluid_sensor% Then Input #fn1, dummy
        Input #fn1, Q, P
        If Q = 0 And P = 0 Then
            Close #fn1
            MsgBox ts$(40), 0, ts(38)   ' "Data Averaging cannot be run on a Bubble Point data file!"/"Data Averaging"
            Begin.Enabled = True
            Begin.SetFocus
            TestFile = -9
            Exit Function
        Else
            Filetype% = 1
        End If
    Else
        Filetype% = 2
    End If
    
    Close #fn1
    TestFile = Filetype%
    
End Function

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    ' Curve.Caption set in Form_load
    Begin.Caption = gpps2("curve", "begin", language$, Begin.Caption)
    set_fontname Begin, default_font
    Help.Caption = gpps2("curve", "help", language$, Help.Caption)
    set_fontname Help, default_font
    cancel.Caption = gpps2("curve", "cancel", language$, cancel.Caption)
    set_fontname cancel, default_font
    AddFile.Caption = gpps2("curve", "addfile", language$, AddFile.Caption)
    set_fontname AddFile, default_font
    Label1.Caption = get_thing("curve", "label1", language$, Label1.Caption, Label1, default_font)
    Label2.Caption = get_thing("curve", "label2", language$, Label2.Caption, Label2, default_font)
    Label9.Caption = get_thing("curve", "label9", language$, Label9.Caption, Label9, default_font)
    Label11.Caption = get_thing("curve", "label11", language$, Label11.Caption, Label11, default_font)
    Label12.Caption = get_thing("curve", "label12", language$, Label12.Caption, Label12, default_font)
    set_fontstuff Picture3, default_font
    
    ' Other text
    ts$(1) = gpps2("curve", "ts1", language$, "Unit")
    ts$(2) = gpps2("curve", "ts2", language$, "The first selected file will determine the type.")
    ts$(3) = gpps2("curve", "ts3", language$, "Data Averager")
    ts$(4) = gpps2("curve", "ts4", language$, "Averaging cannot be done with Pressure Hold Test files.")
    ts$(5) = gpps2("curve", "ts5", language$, "The selected file is a Capillary Flow test data file.")
    ts$(6) = gpps2("curve", "ts6", language$, "The selected file is a Permeability data file.")
    ts$(7) = gpps2("curve", "ts7", language$, "All files must be of the same type.")
    ts$(8) = gpps2("curve", "ts8", language$, "The maximum number of reports is selected.")
    ts$(9) = gpps2("curve", "ts9", language$, "Press BEGIN to continue.")
    ts$(10) = gpps2("curve", "ts10", language$, "You must select at least 1 file.")
    ts$(11) = gpps2("curve", "ts11", language$, "The minimum value for pressure interval")
    ts$(12) = gpps2("curve", "ts12", language$, "Curve Fit")
    ts$(13) = gpps2("curve", "ts13", language$, "Gas Permeameter")
    ts$(14) = gpps2("curve", "ts14", language$, "CFA")
    ts$(15) = gpps2("curve", "ts15", language$, "Integrity Tester")
    ts$(16) = gpps2("curve", "ts16", language$, "Bubble Point Tester")
    ts$(17) = gpps2("curve", "ts17", language$, "Capillary Flow Porometer")
    ts$(18) = gpps2("curve", "ts18", language$, "Data File")
    ts$(19) = gpps2("curve", "ts19", language$, "Warning:  Fluid Sensor data will not be saved to new file")
    ts$(20) = gpps2("curve", "ts20", language$, "Curve Fitting cannot be done on Pressure Hold Test data files.")
    ts$(21) = gpps2("curve", "ts21", language$, "Curve fitting cannot be run on a Bubble Point data file!")
    ts$(22) = gpps2("curve", "ts22", language$, "Cannot curve fit file")
    ts$(23) = gpps2("curve", "ts23", language$, "Not enough points in file.")
    ts$(24) = gpps2("curve", "ts24", language$, "Data Error")
    ts$(25) = gpps2("curve", "ts25", language$, "wet data points and")
    ts$(26) = gpps2("curve", "ts26", language$, "dry data points in")
    ts$(27) = gpps2("curve", "ts27", language$, "File")
    ts$(28) = gpps2("curve", "ts28", language$, "Data")
    ts$(29) = gpps2("curve", "ts29", language$, "data points in file")
    ts$(30) = gpps2("curve", "ts30", language$, "wet")
    ts$(31) = gpps2("curve", "ts31", language$, "dry")
    ts$(32) = gpps2("curve", "ts32", language$, "Percent Complete")
    ts$(33) = gpps2("curve", "ts33", language$, "Curve-Fit Data File SAVE AS")
    ts$(34) = gpps2("curve", "ts34", language$, "Averaged Data File SAVE AS")
    ts$(35) = gpps2("curve", "ts35", language$, "There is a problem with your data file.")
    ts$(36) = gpps2("curve", "ts36", language$, "File Error")
    ts$(37) = gpps2("curve", "ts37", language$, "Reading Files")
    ts$(38) = gpps2("curve", "ts38", language$, "Data Averager")
    ts$(39) = gpps2("curve", "ts39", language$, "Polynomial Curve Fitting")
    ts$(40) = gpps2("curve", "ts40", language$, "Data Averaging cannot be run on a Bubble Point data file!")
    If hhRunAsBurst(current_unit%) Then
        ts$(41) = "Burst Tester Data File"
    Else
        ts$(41) = gpps2("curve", "ts41", language$, "Hydrohead Tester Data File")
    End If
    
End Sub
