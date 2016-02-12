VERSION 5.00
Begin VB.Form BPSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bubble Point Tester Options"
   ClientHeight    =   5040
   ClientLeft      =   3870
   ClientTop       =   3600
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   330
      Left            =   5865
      TabIndex        =   23
      Text            =   "9"
      Top             =   2415
      Width           =   780
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Auto Fill Pump"
      Height          =   255
      Left            =   4230
      TabIndex        =   22
      Top             =   1980
      Width           =   2115
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   330
      Left            =   5895
      TabIndex        =   21
      Text            =   "30"
      Top             =   1380
      Width           =   780
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Enable Post Test Purge"
      Height          =   255
      Left            =   4245
      TabIndex        =   16
      Top             =   600
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   420
      Left            =   5670
      TabIndex        =   13
      Top             =   4455
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4335
      TabIndex        =   12
      Top             =   4455
      Width           =   1200
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   2970
      TabIndex        =   9
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto Wetting"
      Height          =   4185
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   6765
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5235
         TabIndex        =   28
         Top             =   3555
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Reduce flow at pressure target"
         Height          =   420
         Left            =   4140
         TabIndex        =   27
         Top             =   3090
         Width           =   2415
      End
      Begin VB.CheckBox cbPressOnWait 
         Caption         =   "Pressurize on Timeout"
         Height          =   255
         Left            =   4130
         TabIndex        =   26
         Top             =   2680
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   330
         Left            =   5790
         TabIndex        =   20
         Text            =   "400"
         Top             =   810
         Width           =   780
      End
      Begin VB.TextBox Text6 
         Height          =   330
         Left            =   2865
         TabIndex        =   14
         Top             =   2625
         Width           =   1080
      End
      Begin VB.TextBox Text5 
         Height          =   330
         Left            =   2865
         TabIndex        =   11
         Top             =   2235
         Width           =   1080
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   2865
         TabIndex        =   10
         Top             =   1860
         Width           =   1080
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   2865
         TabIndex        =   8
         Top             =   1110
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   2865
         TabIndex        =   7
         Top             =   735
         Width           =   1065
      End
      Begin VB.CheckBox auto_wet_check 
         Caption         =   "Enable Auto Wet Sample Process"
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   375
         Width           =   3555
      End
      Begin VB.Label Label12 
         Caption         =   "Pressure:"
         Height          =   210
         Left            =   4230
         TabIndex        =   29
         Top             =   3630
         Width           =   990
      End
      Begin VB.Label Label8 
         Caption         =   "Pump Position:"
         Height          =   300
         Left            =   4230
         TabIndex        =   25
         Top             =   2295
         Width           =   1395
      End
      Begin VB.Label Label11 
         Caption         =   "Time(Seconds):"
         Height          =   270
         Left            =   4335
         TabIndex        =   19
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label10 
         Caption         =   "Regulator Counts:"
         Height          =   300
         Left            =   4335
         TabIndex        =   18
         Top             =   915
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Bubble Point Timeout (Seconds):"
         Height          =   300
         Left            =   525
         TabIndex        =   15
         Top             =   2670
         Width           =   2370
      End
      Begin VB.Label Label5 
         Caption         =   "Rotating Motor Speed(0-255):"
         Height          =   300
         Left            =   525
         TabIndex        =   6
         Top             =   2310
         Width           =   2370
      End
      Begin VB.Label Label4 
         Caption         =   "Fill Height(%):"
         Height          =   300
         Left            =   510
         TabIndex        =   5
         Top             =   1935
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Drain Time(Seconds):"
         Height          =   300
         Left            =   495
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Soak Time(Seconds):"
         Height          =   300
         Left            =   480
         TabIndex        =   3
         Top             =   1170
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Fill Time(Seconds):"
         Height          =   300
         Left            =   480
         TabIndex        =   2
         Top             =   765
         Width           =   1935
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Regulator Counts:"
      Height          =   300
      Left            =   4410
      TabIndex        =   24
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Label Label9 
      Caption         =   "Regulator Counts:"
      Height          =   240
      Left            =   4440
      TabIndex        =   17
      Top             =   2355
      Width           =   1410
   End
End
Attribute VB_Name = "BPSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub auto_wet_check_Click()
If auto_wet_check.value = 0 Then
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
Else
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
End If

End Sub


Private Sub Check1_Click()
If Check1.value = 1 Then
    Text7.Enabled = True
Else
    Text7.Enabled = False
End If
End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then
    Text9.Enabled = True
    Text10.Enabled = True
Else
    Text10.Enabled = False
    Text9.Enabled = False
End If
End Sub

Private Sub Check3_Click()
If Check3.value = 1 Then
    Text8.Enabled = True
Else
    Text8.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If auto_wet_check.value = 1 Then
    prefsForm.auto_wet_check.value = 1
    prefsForm.time1.Text = Text1.Text
    prefsForm.time4.Text = "0"
    prefsForm.optWetVolume.value = False
    prefsForm.time2.Text = Text2.Text
    prefsForm.time3.Text = Text3.Text
    auto_wet_fill_height = Text4.Text
    auto_wet_rotating_speed = CInt(Text5.Text)
    WPPS Curr_U$, "AutoFillHeight", auto_wet_fill_height, IFile$
    WPPS Curr_U$, "MinBPTime", Text6.Text, IFile$
    bubWaitTime = CInt(Text6.Text)
    WPPS Curr_U$, "MinBPTime", Text6.Text, IFile$
    bubPressOnWait = cbPressOnWait.value
    If cbPressOnWait.value = 1 Then
        'write out here
        WPPS Curr_U$, "BPPressOnWait", "Y", IFile$
    Else
        'writeoutfalse ere
        WPPS Curr_U$, "BPPressOnWait", "N", IFile$
    End If
    prefsForm.txtPumpSpeed.Text = Text5.Text
Else
    prefsForm.auto_wet_check.value = 0
   ' If geoPoreValve = True Then
        WPPS Curr_U$, "MinBPTime", Text6.Text, IFile$
        bubWaitTime = CInt(Text6.Text)
    'End If
End If
If Check2.value = 1 Then
    'save shit
    WPPS Curr_U$, "BPPostPurge", "Y", IFile$
    BPPostPurge = True
    WPPS Curr_U$, "BPPostPurgeCounts", Text9.Text, IFile$
    BPPostPurgeCounts% = CInt(Text9.Text)
    WPPS Curr_U$, "BPPostPurgeDuration", Text10.Text, IFile$
    BPPostPurgeDuration = CInt(Text10.Text)
Else
    'don't save shit
    WPPS Curr_U$, "BPPostPurge", "N", IFile$
    BPPostPurge = False
End If
'i = GPPS("Capstuff", "PneumaticMotor", "N", Ret$, 255, CSFile$)
'If (Left$(UCase$(LTrim$(Ret$)), 1) = "Y") Then
'    PneumaticMotor = True
'    i = GPPS("Capstuff", "PneumaticMotorVNum", "1", Ret$, 255, CSFile$)
'    'pnumValve% = CLng(nulltrim(Ret$))
'    pnumValve% = RonValvePosition(CLng(nulltrim(Ret$)) - 1)
'End If
If Check1.value = 1 Then
    Call WPPS(Curr_U$, "PneumaticMotor", "Y", IFile$)
    PneumaticMotor = True
    Call WPPS(Curr_U$, "PneumaticMotorVNum", Text7.Text, IFile$)
    pnumValve% = CInt(Text7.Text)
Else
    Call WPPS(Curr_U$, "PneumaticMotor", "N", IFile$)
    PneumaticMotor = False
End If

If Check3.value = 1 Then
    ReduceFlowAtTarget = True
    ReduceFlowPressureTarget = CSng(Text8.Text)
    WPPS Curr_U$, "ReduceFlowAtTarget", "Y", IFile$
    WPPS Curr_U$, "ReduceFlowPressureTarget", CStr(ReduceFlowPressureTarget), IFile$
Else
    ReduceFlowAtTarget = False
    WPPS Curr_U$, "ReduceFlowAtTarget", "N", IFile$
End If
    
Unload Me
End Sub

Private Sub Form_Load()
If auto_wet_enable = True Then
    auto_wet_check.value = 1
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
Else
    auto_wet_check.value = 0
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
End If

Text1.Text = Format$(auto_wet_wet_time)
Text2.Text = Format$(auto_wet_soak_time)
Text3.Text = Format$(auto_wet_drain_time)
Text4.Text = Format$(auto_wet_fill_height)
prefsForm.optWetVolume.value = False
Text5.Text = Format$(auto_wet_pump_speed)
Text6.Text = Format$(bubWaitTime)
cbPressOnWait.value = IIf(bubPressOnWait, 1, 0)

Text8.Text = Format$(ReduceFlowPressureTarget)
If ReduceFlowAtTarget = True Then
    Check3.value = 1
    Text8.Enabled = True
End If

'GPPS "Capstuff", "ReduceFlowAtTarget", "N", Ret$, 2, CSFile$
'    ReduceFlowAtTarget = IIf(Left(Ret$, 1) = "Y", True, False)
'    ReduceFlowPressureTarget = val(gpps2("Capstuff", "ReduceFlowPressureTarget", CSFile$, "0"))

    Text6.Enabled = True

If BPPostPurge = True Then
    Check2.value = 1
Else
    Check2.value = 0
End If
Text9.Text = str$(BPPostPurgeCounts)
Text10.Text = str(BPPostPurgeDuration)
If PneumaticMotor = True Then
    Check1.value = 1
    Text7.Enabled = True
End If
Text7.Text = str$(pnumValve%)

End Sub


