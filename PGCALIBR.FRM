VERSION 5.00
Begin VB.Form PGCalibrate 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gauge Calibration"
   ClientHeight    =   4230
   ClientLeft      =   1500
   ClientTop       =   1845
   ClientWidth     =   7590
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
   ForeColor       =   &H00FF00FF&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   7590
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10000
      Top             =   1920
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton Exit 
         Appearance      =   0  'Flat
         Caption         =   "Exit"
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Information  Box:"
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6975
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Width           =   6255
         End
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "High Pressure Gauge ( H-Range )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1605
         Width           =   3495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "High Pressure Gauge ( L-Range )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   2085
         Width           =   3495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Low Pressure Gauge ( H-Range )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   2580
         Width           =   3495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Low Pressure Gauge ( L-Range )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   3045
         Width           =   3495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Flow Rate   ( cc/min.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label L 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   3840
         TabIndex        =   8
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label L 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   7
         Top             =   2040
         Width           =   1755
      End
      Begin VB.Label L 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   3840
         TabIndex        =   6
         Top             =   2520
         Width           =   1755
      End
      Begin VB.Label L 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   3840
         TabIndex        =   5
         Top             =   3000
         Width           =   1755
      End
      Begin VB.Label p1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3840
         TabIndex        =   4
         Top             =   3600
         Width           =   1755
      End
   End
End
Attribute VB_Name = "PGCalibrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ### Based on CapWinV6.54.20 on 3/26/97 ###
' Jing makes some changes to fix the bug on gauge calibration for a machine with only one
' 500 PSI pressure gauge
' The changes are made on both CalibratePG and CheckPG in CaliThirdFull:
' "While X5 < NN%" ->> "Do while X5 < NN%"
' "Wend" ->> "End"
' Add one line "If x4 > 17000 Then Exit Do" to exit the do loop when the pressure approaching the
' target pressure

' ### Based on CapWinV6.54.23 on 6/5/97
' Modification in PGCalibrate to properly handle the following case:
' Only one absolute pressure gauge with range 250 PSI and two flow meter with range 30 cc/min and 500L/min.

Dim Gauge_Case As Integer, Error_Flag As Integer
Dim X As Single, Y As Single, r As Single, dx As Single, x1 As Single, y1 As Single
Dim r1 As Single, y3  As Single, y4 As Single

Private Sub CalibratePG()
Dim i%, NN%, x_temp%, old_pres%, Cpres%, new_pres%
Dim CV2Pos As Integer
Dim Temp_X5 As Single
Dim T As Single
Dim P0 As Single
Dim temp_x4 As Integer
Dim k As Single
Dim LowX4 As Integer
    Dim starttimecount As Single
    Dim JudgementLooser As Single
    Dim hfzero As Single
    
    Label3.Caption = "Pressure gauge calibration begins" + vbCrLf$ + "Setting to atmospheric pressure..."
    L(1).ForeColor = &HFFFF00
    If Gauge_Case < 3 Then
       Label2(2).Visible = False
       Label2(3).Visible = False
    End If
    Show 0
    If GaugeCali_Aborted = True Then Exit Sub
'*** setting to atmospheric pressure ***
    For i% = 0 To 1
       Pres% = i%
       testreadxreturnx4 2
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    Next i%
    If x5 < PY2(2) Then
       Move_Valve 10, "O"
       For i% = 2 To 3
          Pres% = i%
          testreadxreturnx4 2
          If x5 < PY2(i%) Then
             L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
          End If
       Next i%
    End If
    Pres% = 0: testreadxreturnx4 2
    If x5 > 17 Then Zero_Reg 'double zero reg. to make sure it reaches atm. when needed
    If GaugeCali_Aborted = True Then Exit Sub
    Zero_Reg
    If GaugeCali_Aborted = True Then Exit Sub
    Move_Valve 0, "C"
    Move_Valve 11, "O"
    Move_Valve 2, "O"
    Pres% = 1: testreadxreturnx4 2: L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    If x5 < PY2(2) Then Move_Valve 10, "O"
    Label3.Caption = "Stabilizing the system..."
    If Gauge_Case > 2 Then
       NN% = 3
    Else
       NN% = 1
    End If
    lflow% = 0: testreadxreturnx4 0
    If x4 >= 23000 Then
       p1.Caption = " >" + Xformat$(x5, "####0.0")
    Else
       p1.Caption = Xformat$(x5, "####0.0")
    End If
    starttimecount = Timer
    JudgementLooser = 0
    While x5 > 0.8 + JudgementLooser
       If (Timer - starttimecount) > 10 Then
          JudgementLooser = JudgementLooser + 0.1
          starttimecount = Timer
       End If
       For i% = 0 To NN%
          Pres% = i%
          testreadxreturnx4 2
          L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
       Next i%
       lflow% = 0: testreadxreturnx4 0
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       If GaugeCali_Aborted = True Then Exit Sub
       DoEvents
    Wend
'   --- close V2 completely ---
    CV2Pos = V2POS
    V2POS = CLIMIT
    Move2V2Pos
    V2POS = CV2Pos
    If GaugeCali_Aborted = True Then Exit Sub
'   *** set atmospheric pressure value ***
'   --- set P0 by LrHg ---
    If Gauge_Case > 2 Then
       NN% = 3
    Else
       NN% = 1
    End If
    For i% = 0 To NN%
       Pres% = i%
       testreadxreturnx4 2
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
       If GaugeCali_Aborted = True Then Exit Sub
    Next i%
    Pres% = 1
    If Gauge_Case = 3 Or Gauge_Case = 4 Then Pres% = 2
    testreadxreturnx4 2
    L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    If x4 < 3 Then
       Error_Flag = 0
       MsgBox "Warning: " + Gauge_Label$(Pres%, Gauge_Case) + vbCrLf$ + "is not calibrated correctly"
    End If
    P0 = x5
    real_atm = P0
    atm_x4(Pres%) = x4
'   --- calibrate HrHg by LrHg, if it is absolute gauge ---
    If GaugeCali_Aborted = True Then Exit Sub
    Pres% = 0
    testreadxreturnx4 2
    L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    atm_x4(0) = x4
    Temp_X5 = PY2(0) - PY1(0)
    PY1(0) = P0 - (x4 - 500) * Temp_X5 / 20000
    PY2(0) = P0 + (20500 - x4) * Temp_X5 / 20000
    testreadxreturnx4 2
    L(Format(Pres%)).ForeColor = &HFFFF& 'yellow
    L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
'   --- calibrate HrLg and LrLg, if they are differential ---
    If Gauge_Case >= 5 Then
       If x5 < PY2(2) Then
          Pres% = 2: Move_Valve 10, "O"
       End If
       testreadxreturnx4 2
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
       If x4 < 3 Then
          Error_Flag = 0
          MsgBox "Warning: " + Gauge_Label$(Pres%, Gauge_Case) + vbCrLf$ + "is not calibrated correctly"
       End If
       Temp_X5 = PY2(2) - PY1(2)
       PY1(2) = P0 - (x4 - 500) * Temp_X5 / 20000
       PY2(2) = P0 + (20500 - x4) * Temp_X5 / 20000
       atm_x4(2) = x4
       testreadxreturnx4 2
       L(Format(Pres%)).ForeColor = &HFFFF& 'yellow
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
       If GaugeCali_Aborted = True Then Exit Sub
       If x5 < PY2(2) Then
          Pres% = 3: Move_Valve 10, "O"
       End If
       testreadxreturnx4 2
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
       If x4 < 3 Then
          Error_Flag = 0
          MsgBox "Warning: " + Gauge_Label$(Pres%, Gauge_Case) + vbCrLf$ + "is not calibrated correctly"
       End If
       Temp_X5 = PY2(3) - PY1(3)
       PY1(3) = P0 - (x4 - 500) * Temp_X5 / 20000
       PY2(3) = P0 + (20500 - x4) * Temp_X5 / 20000
       atm_x4(3) = x4
       testreadxreturnx4 2
       L(Format(Pres%)).ForeColor = &HFFFF& 'yellow
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
       If GaugeCali_Aborted = True Then Exit Sub
'   --- reset P0 by HrLg if it is differential, and calibrate HrHg and LrHg if they are absolute ---
    ElseIf Gauge_Case = 4 Then
       If PY1(1) <= 5 Then
          x_temp% = atm_x4(1) - ((P0 - PY1(1)) / (PY2(1) - PY1(1)) * (PX2(1) - PX1(1)) + PX1(1))
          PX1(1) = PX1(1) + x_temp%
          PX2(1) = PX2(1) + x_temp%
       End If
       atm_x4(3) = 0
    End If
    old_pres% = Pres%
    If GaugeCali_Aborted = True Then Exit Sub
    ' if the drain valve is lower liquid, we want it open for this
    If Drain12 Then
        Move_Valve 11, "O"
    Else
        Move_Valve 11, "C"
    End If
    Move_Valve 2, "C"
'*** calibrate the first full scale ***
    Pres% = 1: testreadxreturnx4 2
    If Gauge_Case = 3 Or Gauge_Case = 4 Then
       If x5 < PY2(2) Then
          Pres% = 2: Move_Valve 10, "O"
       End If
    ElseIf Gauge_Case > 4 Then
       If x5 < PY2(2) Then
          Pres% = 3: Move_Valve 10, "O"
       End If
    End If
    Label3.Caption = "Calibrating " + " " + Gauge_Label$(Pres% - 1, Gauge_Case) + "..."
    'Refresh
    L(Format(Pres%)).ForeColor = &HFFFF00          'blue
    If GaugeCali_Aborted = True Then Exit Sub
    If Gauge_Case < 3 Then
       GoTo CaliThirdFull
    ElseIf (Gauge_Case = 6) Or (Gauge_Case = 8) Then
       NN% = 17
    ElseIf (Gauge_Case = 5) Or (Gauge_Case = 7) Then
       NN% = 7
    Else
       NN% = 20
    End If
    For i% = 1 To NN%
        inc_reg 10
        If GaugeCali_Aborted = True Then Exit Sub
    Next i%
    Cpres% = Pres%
    For i% = 0 To 3
       Pres% = i%
       testreadxreturnx4 2
       If i% <> Cpres% Then
          L(Format(Pres%)).ForeColor = &HFF00&
       Else
          L(Format(Pres%)).ForeColor = &HFFFF00
       End If
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    Next i%
    Pres% = Cpres%
    Move_Valve 0, "O"
    Label3.Caption = "Target pressure is about " + Xformat$(PY2(Pres%) * PCNV, "###.000") + " " + PU$
    If Gauge_Case = 5 Or Gauge_Case = 7 Then
       NN% = 3
    Else
       NN% = 7
    End If
CheckFirstFullScale:
    For i% = 0 To 3
        If i% <> Cpres% Then
           Pres% = i%
           testreadxreturnx4 2
           L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
        End If
    Next i%
    Pres% = Cpres%
    testreadxreturnx4 2
    L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    If x4 < 18500 Then     'before =20500
       If GaugeCali_Aborted = True Then Exit Sub
       lflow% = 0: testreadxreturnx4 0
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       If (Timer - T > 2 And Timer >= T) And x5 < NN% + 2 Then
          T = Timer
          For i% = 1 To 2
              inc_reg 10
          Next i%
       End If
       DoEvents
       GoTo CheckFirstFullScale
    Else
       Move_Valve 0, "C"
       Label3.Caption = "Stabilizing the system... "
       lflow% = 0: testreadxreturnx4 0
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       starttimecount = Timer
       JudgementLooser = 0
       While x5 > 0.8 + JudgementLooser
             If (Timer - starttimecount) > 10 Then
                JudgementLooser = JudgementLooser + 0.1
                starttimecount = Timer
             End If
             For i% = 0 To 3
                Pres% = i%
                testreadxreturnx4 2
                L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
             Next i%
             If GaugeCali_Aborted = True Then Exit Sub
             lflow% = 0: testreadxreturnx4 0
             If x4 >= 23000 Then
                p1.Caption = " >" + Xformat$(x5, "####0.0")
             Else
                p1.Caption = Xformat$(x5, "####0.0")
             End If
             DoEvents
       Wend
    End If
    Pres% = Cpres%
    GoTo CaliPG
CaliSecondFull:
    Pres% = 2: testreadxreturnx4 2
    If x5 < PY2(2) Then
       Move_Valve 10, "O"
    End If
    If Gauge_Case = 4 Then Pres% = 1
    If GaugeCali_Aborted = True Then Exit Sub
    Label3.Caption = "Calibrating " + " " + Gauge_Label$(Pres% - 1, Gauge_Case)
    If GaugeCali_Aborted = True Then Exit Sub
    If Gauge_Case = 4 Then
       NN% = 25
    ElseIf (Gauge_Case = 6) Or (Gauge_Case = 8) Then
       NN% = 15
    ElseIf (Gauge_Case = 5) Or (Gauge_Case = 7) Then
       NN% = 10
    Else
       NN% = 15           'before =7
    End If
    If (Gauge_Case = 5) Or (Gauge_Case = 7) Then
       For i% = 1 To NN%
          inc_reg 10
          If GaugeCali_Aborted = True Then Exit Sub
       Next i%
       Move_Valve 0, "O"
    Else
       For i% = 1 To NN%
           inc_reg 10
          If GaugeCali_Aborted = True Then Exit Sub
       Next i%
       Move_Valve 0, "O"
       If GaugeCali_Aborted = True Then Exit Sub
    End If
    Cpres% = Pres%
    For i% = 0 To 3
       Pres% = i%
       testreadxreturnx4 2
       If i% <> Cpres% Then
          L(Format(Pres%)).ForeColor = &HFF00&
       Else
          L(Format(Pres%)).ForeColor = &HFFFF00
       End If
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    Next i%
    Pres% = Cpres%
    If GaugeCali_Aborted = True Then Exit Sub
'*** calibrate the Second full scale ***
    Label3.Caption = "Calibrating" + " " + Gauge_Label$(Pres% - 1, Gauge_Case)
    Label3.Caption = "Target pressure is about " + Xformat$(PY2(Pres%) * PCNV, "###.000") + " " + PU$
    L(3).Caption = " "
    T = Timer
    If Gauge_Case = 5 Or Gauge_Case = 7 Then
       NN% = 7
    Else
       NN% = 10
    End If
CheckSecondFullScale:
    For i% = 0 To 2
           Pres% = i%
           testreadxreturnx4 2
           L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    Next i%
    Pres% = Cpres%
    If x4 < 20000 Then         'before 20500
       If GaugeCali_Aborted = True Then Exit Sub
       lflow% = 0: testreadxreturnx4 0
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       If (Timer - T > 5 And Timer >= T) And x5 < NN% Then
          T = Timer
          For i% = 1 To 2
              inc_reg 10
          Next i%
       End If
       DoEvents
       GoTo CheckSecondFullScale
    Else
       Move_Valve 0, "C"
       Label3.Caption = "Stabilizing the system..."
       T = Timer
       Cpres% = Pres%
       lflow% = 0: testreadxreturnx4 0
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       starttimecount = Timer
       JudgementLooser = 0
       While x5 > 0.8 + JudgementLooser
          If (Timer - starttimecount) > 5 Then
             JudgementLooser = JudgementLooser + 0.2
             starttimecount = Timer
          End If
             For i% = 0 To 2
                Pres% = i%
                testreadxreturnx4 2
                L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
             Next i%
             If GaugeCali_Aborted = True Then Exit Sub
             lflow% = 0: testreadxreturnx4 0
             If x4 >= 23000 Then
                p1.Caption = " >" + Xformat$(x5, "####0.0")
             Else
                p1.Caption = Xformat$(x5, "####0.0")
             End If
             DoEvents
       Wend
    End If
    Pres% = Cpres%
    GoTo CaliPG
    Move_Valve 10, "C"
    L(Format(Pres%)).Caption = " "
'  --- calibrate LrHg at HrLg's full range ---
CaliThirdFull:
    Pres% = 1
    Label3.Caption = "Calibrating" + " " + Gauge_Label$(Pres% - 1, Gauge_Case) + "..."
    If GaugeCali_Aborted = True Then Exit Sub
    If Gauge_Case = 1 Then
       NN% = 10: Cpres% = 1
    ElseIf Gauge_Case = 2 Then
       NN% = 45: Cpres% = 1
    ElseIf Gauge_Case < 7 Then
       NN% = 10
    Else
       NN% = 10      'before =35
    End If
    For i% = 1 To NN%
       If GaugeCali_Aborted = True Then Exit Sub
       inc_reg 10
    Next i%
    HFLOW% = 1: testreadxreturnx4 1
    hfzero = x5
    If (Gauge_Case = 3) Or (Gauge_Case = 5) Or (Gauge_Case = 6) Then
       NN% = 150
    ElseIf Gauge_Case = 1 Then
       NN% = 50
    ElseIf Gauge_Case = 8 Or Gauge_Case = 2 Then
       If PY2(1) < 50 Then
          NN% = 150
       ElseIf PY2(1) < 110 Then
          NN% = 800
       Else
          NN% = 1500
       End If
    Else
       NN% = 100
    End If
    L(2).Caption = " "
    starttimecount = Timer
    x5 = 0
    Do While x5 < NN%
        If x5 < 50 Then
           If (Timer - starttimecount) > 1 Then
              inc_reg 10
              inc_reg 10
              Pulse_V2 0
              starttimecount = Timer
           End If
        Else
           If (Timer - starttimecount) > 2 Then
              inc_reg 10
              Pulse_V2 0
              'Send_RS232 ("J" + Chr$(6) + "IBJ" + Chr$(12))
              starttimecount = Timer
           End If
        End If
        For i% = 0 To 1
           Pres% = i%
           testreadxreturnx4 2
           If i% <> Cpres% Then
              L(Format(Pres%)).ForeColor = &HFF00&
           Else
              L(Format(Pres%)).ForeColor = &HFFFF00
           End If
           L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
        Next i%
        If x4 > 17000 Then Exit Do
        HFLOW% = 1: testreadxreturnx4 1
        x5 = x5 - hfzero
        If x4 >= 23000 Then
           p1.Caption = " >" + Xformat$(x5, "####0.0")
        Else
           p1.Caption = Xformat$(x5, "####0.0")
        End If
        If GaugeCali_Aborted = True Then Exit Sub
    Loop
    starttimecount = Timer
    V2POS = CV2Pos
    Cpres% = Pres%
    Pres% = Cpres%
    Label3.Caption = "Target pressure is about " + Xformat$(PY2(Pres%) * PCNV, "###.000") + " " + PU$
CheckThirdFullScale:
    If Gauge_Case <> 2 Then
       Move_Valve 10, "C"
    Else
       Move_Valve 0, "O"
    End If
    For i% = 0 To 1
        If i% <> Pres% Then
           Pres% = i%
           testreadxreturnx4 2
           L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
        End If
    Next i%
    Pres% = Cpres%
    testreadxreturnx4 2

    If x4 < 19500 Then
       
       If GaugeCali_Aborted = True Then Exit Sub
       If Gauge_Case = 2 Then
          lflow% = 0: testreadxreturnx4 0
       Else
          HFLOW% = 1: testreadxreturnx4 1
          x5 = x5 - hfzero
       End If
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       If (Timer - starttimecount) > 2 And x5 < 50 Then ' yong
          inc_reg 10
          inc_reg 10
          If Gauge_Case <> 2 Then Pulse_V2 0
          starttimecount = Timer
       End If
       DoEvents
       If GaugeCali_Aborted = True Then Exit Sub
       GoTo CheckThirdFullScale
    Else
       Move_Valve 0, "C"
       CV2Pos = V2POS
       V2POS = CLIMIT
       Move2V2Pos
       V2POS = CV2Pos
       Label3.Caption = "Stabilizing the system..."
       If Gauge_Case = 2 Then
          lflow% = 0: testreadxreturnx4 0
       Else
          HFLOW% = 1: testreadxreturnx4 1
          x5 = x5 - hfzero
       End If
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       starttimecount = Timer
       JudgementLooser = 0
       While x5 > 13 + JudgementLooser
             If (Timer - starttimecount) > 5 Then
                JudgementLooser = JudgementLooser + 4
                starttimecount = Timer
             End If
             For i% = 0 To 1
                Pres% = i%
                testreadxreturnx4 2
                L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
             Next i%
             If GaugeCali_Aborted = True Then Exit Sub
             DoEvents
             If Gauge_Case = 2 Then
                lflow% = 0: testreadxreturnx4 0
             Else
                HFLOW% = 1: testreadxreturnx4 1
                x5 = x5 - hfzero
             End If
             If x4 >= 23000 Then
                p1.Caption = " >" + Xformat$(x5, "####0.0")
             Else
                p1.Caption = Xformat$(x5, "####0.0")
             End If
       Wend
    End If
CaliPG:
    old_pres% = Pres%: LowX4 = x4
    new_pres% = Pres% - 1
    If Gauge_Case = 3 Then new_pres% = Pres% - 2
    Pres% = new_pres%

    x4 = raw_reading(2): temp_x4 = x4
    temp_x4 = x4
    Pres% = old_pres%
    LowX4 = raw_reading(2)
    Pres% = new_pres%
    x4 = temp_x4

    If LowX4 < 23000 And x4 < 20500 Then
       Temp_X5 = (LowX4 - PX1(old_pres%)) / (PX2(old_pres%) - PX1(old_pres%)) * (PY2(old_pres%) - PY1(old_pres%)) + PY1(old_pres%)
       PX1(Pres%) = 500: PX2(Pres%) = 20500
       PY1(Pres%) = (500 - atm_x4(Pres%)) / (x4 - atm_x4(Pres%)) * (Temp_X5 - real_atm) + real_atm
       PY2(Pres%) = (20500 - atm_x4(Pres%)) / (x4 - atm_x4(Pres%)) * (Temp_X5 - real_atm) + real_atm
       X1FIRST = x4
    End If
    testreadxreturnx4 2
    L(Format(Pres%)).ForeColor = &HFFFF&
    L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    'max_count_p = x4
    Move_Valve 10, "C"
    If Gauge_Case > 3 And Pres% = 2 Then GoTo CaliSecondFull
    If Gauge_Case > 4 And Pres% = 1 Then GoTo CaliThirdFull
DoneCali:
    If GaugeCali_Aborted = True Then Exit Sub
    Label3.Caption = "Calibration is done." + vbCrLf$ + "Reset to atmospheric pressure..."
    'Refresh
'*** reset to atmospheric pressure ***
    Pres% = 1
    Zero_Reg
    Move_Valve 0, "C"
    Move_Valve 11, "O"
    Move_Valve 2, "O"
    T = Timer
    Cpres% = Pres%
    While (Timer - T) < 8 And Timer >= T          '8 SECOND DELAY
          If Gauge_Case < 3 Then
             NN% = 1
          Else
             NN% = 3
          End If
          For i% = 0 To NN%
          Pres% = i%
          testreadxreturnx4 2
          If x5 < PY2(2) Then
             Move_Valve 10, "O"
          End If
          L(Format(Pres%)).ForeColor = &HFF00&
          L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
          Next i%
          If GaugeCali_Aborted = True Then Exit Sub
          DoEvents
    Wend
    If Drain12 Then
        Move_Valve 11, "O"
    Else
        Move_Valve 11, "C"
    End If
    Move_Valve 2, "C"
End Sub

Private Sub CheckPG()
Dim old_pres%, NN%, i%, Cpres%, Ret$
Dim Temp_X5 As Single
Dim Rel_Error As Single
Dim CV2Pos As Integer
Dim T As Single
Dim r As Long
Dim hfzero As Single
    Dim starttimecount As Single
    Dim JudgementLooser As Single

    '*** increase pressure and read pressure gauges to make sure calibrate correctly ***
    Label3.Caption = " Checking pressure gauge reading..." + vbCrLf$ + " At Atmospheric Pressure..."
    Pres% = 1
    If Gauge_Case = 3 Or Gauge_Case = 4 Then
       Pres% = 2
    End If
    old_pres% = Pres%
    Zero_Reg
    If GaugeCali_Aborted = True Then Exit Sub
    Move_Valve 0, "C"
    Move_Valve 11, "O"
    Move_Valve 2, "O"
    lflow% = 0: testreadxreturnx4 0
    If x4 >= 23000 Then
       p1.Caption = " >" + Xformat$(x5, "####0.0")
    Else
       p1.Caption = Xformat$(x5, "####0.0")
    End If
    If x5 < PY2(2) Then Move_Valve 10, "O"
    If Gauge_Case > 2 Then
       NN% = 3
    Else
       NN% = 1
    End If
    starttimecount = Timer
    JudgementLooser = 0
    While x5 > 0.8 + JudgementLooser
          If (Timer - starttimecount) > 10 Then
             JudgementLooser = JudgementLooser + 0.1
             starttimecount = Timer
          End If
          For i% = 0 To NN%
             Pres% = i%
             testreadxreturnx4 2
             L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
          Next i%
          lflow% = 0: testreadxreturnx4 0
          If x4 >= 23000 Then
             p1.Caption = " >" + Xformat$(x5, "####0.0")
          Else
             p1.Caption = Xformat$(x5, "####0.0")
          End If
          If GaugeCali_Aborted = True Then Exit Sub
          DoEvents
    Wend
    Pres% = old_pres%
    testreadxreturnx4 2
    Temp_X5 = x5
    L(Format(Pres%)).ForeColor = &HFFFF00    'blue
    NN% = Pres% + 1
    If Gauge_Case = 3 Then
       NN% = Pres%
    ElseIf Gauge_Case > 4 Then
       NN% = Pres% + 3
    End If
    For i% = 1 To NN%
        Pres% = i% - 1
        If Pres% = 1 And Gauge_Case = 3 Then Pres% = 2
        If Pres% > 1 And x5 < PY2(2) Then Move_Valve 10, "O"
        If Pres% <> old_pres% Then
           testreadxreturnx4 2
           Rel_Error = Abs(x5 - Temp_X5) / Temp_X5
           L(Format(Pres%)).ForeColor = &HFFFF&
           L(Format(Pres%)).Caption = Xformat$(x5 - Temp_X5, "####0.000")
           If Rel_Error > 0.003 Then
              Error_Flag = 10
              MsgBox "Warning:" + " " + Gauge_Label$(Pres%, Gauge_Case) + " " + "Calibration failed"
           End If
        End If
        If GaugeCali_Aborted = True Then Exit Sub
        DoEvents
    Next i%
    If Drain12 Then
        Move_Valve 11, "O"
    Else
        Move_Valve 11, "C"
    End If
    Move_Valve 2, "C"
    Label3.Caption = "Reaching " + " " + Gauge_Label$(Pres%, Gauge_Case) + "'s Full Scale..."
ReachFirstFull:
    Pres% = 1: testreadxreturnx4 2
    If Gauge_Case = 3 Or Gauge_Case = 4 Then
       If x5 < PY2(2) Then
          Pres% = 2: Move_Valve 10, "O"
       End If
    ElseIf Gauge_Case > 4 Then
       If x5 < PY2(2) Then
          Pres% = 3: Move_Valve 10, "O"
       End If
    End If
    If GaugeCali_Aborted = True Then Exit Sub
    If Gauge_Case < 3 Then
       GoTo ReachThirdFull
    ElseIf (Gauge_Case = 6) Or (Gauge_Case = 8) Then
       NN% = 17
    ElseIf (Gauge_Case = 5) Or (Gauge_Case = 7) Then
       NN% = 7
    Else
       NN% = 20
    End If
    For i% = 1 To NN%
        inc_reg 10
        If GaugeCali_Aborted = True Then Exit Sub
    Next i%
    Cpres% = Pres%
    For i% = 0 To 3
       Pres% = i%
       testreadxreturnx4 2
       If i% <> Cpres% Then
          L(Format(Pres%)).ForeColor = &HFF00&
       Else
          L(Format(Pres%)).ForeColor = &HFFFF00
       End If
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    Next i%
    Pres% = Cpres%
    Move_Valve 0, "O"
    Label3.Caption = "Target pressure is about " + Xformat$(PY2(Pres%) * PCNV, "###.000") + " " + PU$
    'Cpres% = pres%
    If Gauge_Case = 5 Or Gauge_Case = 7 Then
       NN% = 3
    Else
       NN% = 7
    End If
DoubleCheckFirstFullScale:
    For i% = 0 To 3
        If i% <> Pres% Then
           Pres% = i%
           testreadxreturnx4 2
           L(Format(Pres%)).ForeColor = &HFF00&
           L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
        End If
    Next i%
    Pres% = Cpres%
    testreadxreturnx4 2
    L(Format(Pres%)).ForeColor = &HFF00&
    L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    If x4 < 18500 Then       'before =20500
       If GaugeCali_Aborted = True Then Exit Sub
       lflow% = 0: testreadxreturnx4 0
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       If (Timer - T > 2 And Timer >= T) And x5 < NN% + 2 Then
          T = Timer
          For i% = 1 To 2
              inc_reg 10
          Next i%
       End If
       DoEvents
       GoTo DoubleCheckFirstFullScale
    Else
       Move_Valve 0, "C"
       Label3.Caption = "Stabilizing the system... "
       lflow% = 0: testreadxreturnx4 0
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       starttimecount = Timer
       JudgementLooser = 0
       While x5 > 0.8 + JudgementLooser
             If (Timer - starttimecount) > 10 Then
                JudgementLooser = JudgementLooser + 0.1
                starttimecount = Timer
             End If
             For i% = 0 To 3
                Pres% = i%
                testreadxreturnx4 2
                L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
             Next i%
             If GaugeCali_Aborted = True Then Exit Sub
             lflow% = 0: testreadxreturnx4 0
             If x4 >= 23000 Then
                p1.Caption = " >" + Xformat$(x5, "####0.0")
             Else
                p1.Caption = Xformat$(x5, "####0.0")
             End If
             DoEvents
       Wend
    End If
    Pres% = Cpres%
    GoTo ReadPG
ReachSecondFull:
    Pres% = 2: testreadxreturnx4 2
    If x5 < PY2(2) Then
       Move_Valve 10, "O"
    End If
    If Gauge_Case = 4 Then Pres% = 1
    If GaugeCali_Aborted = True Then Exit Sub
    Label3.Caption = "Reaching " + " " + Gauge_Label$(Pres%, Gauge_Case) + "'s Full Scale..."
    If Gauge_Case = 4 Then
       NN% = 25
    ElseIf (Gauge_Case = 6) Or (Gauge_Case = 8) Then
       NN% = 15
    ElseIf (Gauge_Case = 5) Or (Gauge_Case = 7) Then
       NN% = 10
    Else
       NN% = 15                'before =7
    End If
    If (Gauge_Case = 5) Or (Gauge_Case = 7) Then
       For i% = 1 To NN%
          inc_reg 10
          If GaugeCali_Aborted = True Then Exit Sub
       Next i%
       Move_Valve 0, "O"
    Else
       For i% = 1 To NN%
          inc_reg 10
          If GaugeCali_Aborted = True Then Exit Sub
       Next i%
       Move_Valve 0, "O"
       If GaugeCali_Aborted = True Then Exit Sub
    End If
    Cpres% = Pres%
    For i% = 0 To 3
       Pres% = i%
       testreadxreturnx4 2
       L(Format(Pres%)).ForeColor = &HFF00&
       L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    Next i%
    Pres% = Cpres%
    If GaugeCali_Aborted = True Then Exit Sub
    Label3.Caption = "Target pressure is about " + Xformat$(PY2(Pres%) * PCNV, "###.000") + " " + PU$
    L(3).Caption = " "
    T = Timer
    If Gauge_Case = 5 Or Gauge_Case = 7 Then
       NN% = 7
    Else
       NN% = 10
    End If
DoubleCheckSecondFullScale:
    For i% = 0 To 2
           Pres% = i%
           testreadxreturnx4 2
           L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
    Next i%
    Pres% = Cpres%
    If x4 < 20000 Then  'before =20500
       If GaugeCali_Aborted = True Then Exit Sub
       lflow% = 0: testreadxreturnx4 0
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       If (Timer - T > 5 And Timer >= T) And x5 < NN% Then
          T = Timer
          For i% = 1 To 2
              inc_reg 10
          Next i%
       End If
       DoEvents
       GoTo DoubleCheckSecondFullScale
    Else
       Move_Valve 0, "C"
       Label3.Caption = "Stabilizing the system..."
       T = Timer
       Cpres% = Pres%
       lflow% = 0: testreadxreturnx4 0
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       starttimecount = Timer
       JudgementLooser = 0
       While x5 > 0.8 + JudgementLooser
          If (Timer - starttimecount) > 5 Then
             JudgementLooser = JudgementLooser + 0.2
             starttimecount = Timer
          End If
             For i% = 0 To 2
                Pres% = i%
                testreadxreturnx4 2
                L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
             Next i%
             If GaugeCali_Aborted = True Then Exit Sub
             lflow% = 0: testreadxreturnx4 0
             If x4 >= 23000 Then
                p1.Caption = " >" + Xformat$(x5, "####0.0")
             Else
                p1.Caption = Xformat$(x5, "####0.0")
             End If
             DoEvents
       Wend
    End If
    Pres% = Cpres%
    GoTo ReadPG
    Move_Valve 10, "C"
    L(Format(Pres%)).Caption = " "
ReachThirdFull:
    Pres% = 1
    Label3.Caption = "checking at" + " " + Gauge_Label$(Pres%, Gauge_Case) + "'s Full Scale" + "..."
    If GaugeCali_Aborted = True Then Exit Sub
    If Gauge_Case = 1 Then
       NN% = 10: Cpres% = 1
    ElseIf Gauge_Case = 2 Then
       NN% = 45: Cpres% = 1
    ElseIf Gauge_Case < 7 Then
       NN% = 10
    Else
       NN% = 10                 'before =35
    End If
    For i% = 1 To NN%
        If GaugeCali_Aborted = True Then Exit Sub
        inc_reg 10
    Next i%
    HFLOW% = 1: testreadxreturnx4 1
    hfzero = x5
    If (Gauge_Case = 1) Or (Gauge_Case = 3) Or (Gauge_Case = 5) Or (Gauge_Case = 6) Then
       NN% = 150
    ElseIf Gauge_Case = 8 Or Gauge_Case = 2 Then
       If PY2(1) < 50 Then
          NN% = 150
       ElseIf PY2(1) < 110 Then
          NN% = 800
       Else
          NN% = 1500
       End If
    Else
       NN% = 100
    End If
    L(2).Caption = " "
    starttimecount = Timer
    x5 = 0
    Do While x5 < NN%
        If x5 < 50 Then
           If (Timer - starttimecount) > 1 Then
              inc_reg 10
              inc_reg 10
              Pulse_V2 0
              starttimecount = Timer
           End If
        Else
           If (Timer - starttimecount) > 2 Then
              inc_reg 10
              Pulse_V2 0
              'Send_RS232 ("J" + Chr$(6) + "IBJ" + Chr$(12))
              starttimecount = Timer
           End If
        End If
        For i% = 0 To 1
           Pres% = i%
           testreadxreturnx4 2
           If i% <> Cpres% Then
              L(Format(Pres%)).ForeColor = &HFF00&
           Else
              L(Format(Pres%)).ForeColor = &HFFFF00
           End If
           L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
        Next i%
        If x4 > 17000 Then Exit Do
        HFLOW% = 1: testreadxreturnx4 1
        x5 = x5 - hfzero
        If x4 >= 23000 Then
           p1.Caption = " >" + Xformat$(x5, "####0.0")
        Else
           p1.Caption = Xformat$(x5, "####0.0")
        End If
        If GaugeCali_Aborted = True Then Exit Sub
    Loop
    starttimecount = Timer
    V2POS = CV2Pos
    L(2).Caption = " "
    Cpres% = Pres%
    Pres% = Cpres%
    Label3.Caption = "Target pressure is about " + Xformat$(PY2(Pres%) * PCNV, "###.000") + " " + PU$
    L(2).Caption = " "
DoubleCheckThirdFullScale:
    If Gauge_Case <> 2 Then
       Move_Valve 10, "C"
    Else
       Move_Valve 0, "O"
    End If
    For i% = 0 To 1
        If i% <> Pres% Then
           Pres% = i%
           testreadxreturnx4 2
           L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
        End If
    Next i%
    Pres% = Cpres%
    testreadxreturnx4 2
    
    If x4 < 19500 Then
       If GaugeCali_Aborted = True Then Exit Sub
       If Gauge_Case = 2 Then
          lflow% = 0: testreadxreturnx4 0
       Else
          HFLOW% = 1: testreadxreturnx4 1
          x5 = x5 - hfzero
       End If
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       If (Timer - starttimecount) > 2 And x5 < 50 Then ' yong
          inc_reg 10
          inc_reg 10
          If Gauge_Case <> 2 Then Pulse_V2 0
          starttimecount = Timer
       End If
       DoEvents
       If GaugeCali_Aborted = True Then Exit Sub
       GoTo DoubleCheckThirdFullScale
    Else
       Move_Valve 0, "C"
       CV2Pos = V2POS
       V2POS = CLIMIT
       Move2V2Pos
       V2POS = CV2Pos
       Label3.Caption = "Stabilizing the system..."
       If Gauge_Case = 2 Then
          lflow% = 0: testreadxreturnx4 0
       Else
          HFLOW% = 1: testreadxreturnx4 1
          x5 = x5 - hfzero
       End If
       If x4 >= 23000 Then
          p1.Caption = " >" + Xformat$(x5, "####0.0")
       Else
          p1.Caption = Xformat$(x5, "####0.0")
       End If
       starttimecount = Timer
       JudgementLooser = 0
       While x5 > 13 + JudgementLooser
             If (Timer - starttimecount) > 5 Then
                JudgementLooser = JudgementLooser + 4
                starttimecount = Timer
             End If
             For i% = 0 To 1
                Pres% = i%
                testreadxreturnx4 2
                L(Format(Pres%)).Caption = Xformat$(x5 * PCNV, "#####.000")
             Next i%
             If GaugeCali_Aborted = True Then Exit Sub
             DoEvents
             If Gauge_Case = 2 Then
                lflow% = 0: testreadxreturnx4 0
             Else
                HFLOW% = 1: testreadxreturnx4 1
                x5 = x5 - hfzero
             End If
             If x4 >= 23000 Then
                p1.Caption = " >" + Xformat$(x5, "####0.0")
             Else
                p1.Caption = Xformat$(x5, "####0.0")
             End If
       Wend
    End If
ReadPG:
    old_pres% = Pres%
    testreadxreturnx4 2
    L(Format(Pres%)).ForeColor = &HFFFF00
    Temp_X5 = x5
    NN% = Pres% + 1
    If Gauge_Case = 3 Then NN% = Pres%
    For i% = 1 To NN%
        Pres% = i% - 1
        If Pres% = 1 And Gauge_Case = 3 Then Pres% = 2
        If Pres% > 1 And x5 < PY2(2) Then Move_Valve 10, "O"
        If Pres% <> old_pres% Then
           testreadxreturnx4 2
           Rel_Error = Abs(x5 - Temp_X5) / Temp_X5
           L(Format(Pres%)).ForeColor = &HFFFF&
           L(Format(Pres%)).Caption = Xformat$(x5 - Temp_X5, "####0.000")
           If Rel_Error > 0.002 Then
              Error_Flag = 10
              MsgBox "Warning:" + " " + Gauge_Label$(Pres%, Gauge_Case) + " " + "Calibration failed"
           End If
        End If
        DoEvents
    Next i%
    If Gauge_Case < 4 Then
       GoTo DoneDoubleCheck
    ElseIf Gauge_Case = 4 Then
       If Pres% = 2 Then
          GoTo ReachSecondFull
       Else
          GoTo DoneDoubleCheck
       End If
    Else
       If Pres% = 3 Then
          GoTo ReachSecondFull
       ElseIf Pres% = 2 Then
          GoTo ReachThirdFull
       ElseIf Pres% = 1 Then
          T = Timer
          While (Timer - T) < 5 And Timer >= T          '3 SECOND DELAY
                ' only let them abort every 0.1 seconds
                waitseconds 0.1
                If GaugeCali_Aborted = True Then Exit Sub
                DoEvents
          Wend
          GoTo DoneDoubleCheck
       End If
    End If
DoneDoubleCheck:
    If Error_Flag >= 10 Then
      Label3.Caption = " Automated calibration procedure failed at"
      If Error_Flag = 10 Then Label3.Caption = "atmospheric pressure"
      If Error_Flag = 11 Then Label3.Caption = "Low Pressure (low range)'s full scale"
      If Error_Flag = 12 Then Label3.Caption = "Low Pressure (high range)'s full scale"
      If Error_Flag = 13 Then Label3.Caption = "High Pressure (low range)'s full scale"
      If Error_Flag = 14 Then Label3.Caption = "Pressure (low range)'s full scale"
   ElseIf Error_Flag >= 0 Then
      Label3.Caption = " The pressure gauge need to be physically calibrated"
   Else
      WPPS "Capstuff", "PX1_0", str$(PX1(0)), CSFile$
      WPPS "Capstuff", "PX2_0", str$(PX2(0)), CSFile$
      WPPS "Capstuff", "PY1_0", str$(PY1(0)), CSFile$
      WPPS "Capstuff", "PY2_0", str$(PY2(0)), CSFile$
      WPPS "Capstuff", "PX1_1", str$(PX1(1)), CSFile$
      WPPS "Capstuff", "PX2_1", str$(PX2(1)), CSFile$
      WPPS "Capstuff", "PY1_1", str$(PY1(1)), CSFile$
      WPPS "Capstuff", "PY2_1", str$(PY2(1)), CSFile$
      If ExtraPG Then
         WPPS "Capstuff", "PX1_2", str$(PX1(2)), CSFile$
         WPPS "Capstuff", "PX2_2", str$(PX2(2)), CSFile$
         WPPS "Capstuff", "PY1_2", str$(PY1(2)), CSFile$
         WPPS "Capstuff", "PY2_2", str$(PY2(2)), CSFile$
         WPPS "Capstuff", "PX1_3", str$(PX1(3)), CSFile$
         WPPS "Capstuff", "PX2_3", str$(PX2(3)), CSFile$
         WPPS "Capstuff", "PY1_3", str$(PY1(3)), CSFile$
         WPPS "Capstuff", "PY2_3", str$(PY2(3)), CSFile$
      End If
   End If
    Pres% = 0: testreadxreturnx4 2
    If x5 > 17 Then Zero_Reg 'double zero reg. to make sure it reaches atm. when needed
    Zero_Reg
    Move_Valve 0, "C"
    Move_Valve 11, "O"
    Move_Valve 2, "O"
End Sub

Private Sub Exit_Click()
Dim numdata&
Dim r As Long
    numdata& = 0
    r = WinHelp(hwnd, HelpFile$, Help_Quit, ByVal numdata&)
    reply% = MsgBox("Exit calibration routine ?", 17, "CapWin")
    If reply% = 1 Then
       GaugeCali_Aborted = True
       'Unload me
    End If
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
    X = ScaleWidth * 8.85 / 10
    Y = ScaleHeight * 3 / 5
    r1 = (Y * 1.3) + 1 - 50
    dx = 0.05
    x1 = r1 + X
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - Unit" + str$(unitnumber)
    End If

    Error_Flag = -1
    top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    MsgBox "Put Non-Porous Disk in Sample Chamber and Seal"
    Get_Gauge_Case
    CalibratePG
    CheckPG
End Sub

Private Sub Form_Unload(cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Function Gauge_Label$(ByVal XVal, YVal)
    If XVal = 0 Then
       If YVal < 3 Then
          Gauge_Label$ = "Pressure (high range)"
       Else
          Gauge_Label$ = "High Pressure (high range)"
       End If
    ElseIf XVal = 1 Then
       If YVal < 3 Then
          Gauge_Label$ = "Pressure (low range)"
       Else
          Gauge_Label$ = "High Pressure (low range)"
       End If
    ElseIf XVal = 2 Then
       If YVal > 2 Then Gauge_Label$ = "Low Pressure (high range)"
    Else
       If YVal > 2 Then Gauge_Label$ = "Low Pressure (low range)"
    End If
End Function

Private Sub Get_Gauge_Case()
    If Not ExtraPG And PY2(0) < 150 Then Gauge_Case = 1   'No extra gauge, full range is less than 7755 torr
    If Not ExtraPG And PY2(0) > 150 Then Gauge_Case = 2   'No extra gauge, full range is larger than 7755 torr
    If ExtraPG And PY2(0) < 150 And PY1(2) < 5 Then Gauge_Case = 3   ' with one extra abs. gauge, the normal gauge's full range is less than 7755 torr
    If ExtraPG And PY2(0) > 150 And PY1(2) < 5 Then Gauge_Case = 4   ' with one extra abs. gauge  the normal gauge's full range is larger than 7755 torr
    If ExtraPG And PY2(0) < 150 And PY1(2) > 10 And (PY2(2) - PY1(2)) < 3 Then Gauge_Case = 5  ' with one extra dif. gauge,full range less than 155 torr, the normal gauge's full range is less than 7755 torr
    If ExtraPG And PY2(0) < 150 And PY1(2) > 10 And (PY2(2) - PY1(2)) > 3 Then Gauge_Case = 6  ' with one extra dif. gauge,full range larger than 155 torr, the normal gauge's full range is less than 7755 torr
    If ExtraPG And PY2(0) > 150 And PY1(2) > 10 And (PY2(2) - PY1(2)) < 3 Then Gauge_Case = 7  ' with one extra dif. gauge,full range less than 155 torr, the normal gauge's full range is larger than 7755 torr
    If ExtraPG And PY2(0) > 150 And PY1(2) > 10 And (PY2(2) - PY1(2)) > 3 Then Gauge_Case = 8  ' with one extra dif. gauge,full range larger than 155 torr, the normal gauge's full range is larger than 7755 torr
End Sub

Private Sub linedemo()
y3 = dx + y3
y4 = y3 - 0.2
x1 = X + r1 * Int(Cos(y4) * 100) / 900
y1 = Y + r1 * Int(Sin(y4) * 100) / 600
DrawWidth = 10
Line (X, Y)-(x1, y1), RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
'Line (x, y + 100)-(x1, y1 + 100), RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
End Sub

Private Sub Note()
'### The calibration routine takes about 5 min. ###
'**** structure of the calibration procedure ***
'Define Gauge_Case as a global variable:
'    = 1: one absolute gauge , range 5000/1000(Torr)
'    = 2: one absolute gauge , range 10000/2000(Torr)
'    = 3: two absolute gauges, one range 5000/1000(Torr); another range 1000/200(Torr)
'    = 4: two absolute gauges, one range 10000/2000(Torr); another range 1000/200(Torr)
'    = 5: two gauges, one absolute range 5000/1000(Torr); another differential range 100/20(Torr)
'    = 6: two gauges, one absolute range 5000/1000(Torr); another differential range 200/40(Torr)
'    = 7: two gauges, one absolute range 10000/2000(Torr); another differential range 100/20(Torr)
'    = 8: two gauges, one absolute range 10000/2000(Torr); another differential range 200/40(Torr)
'In case 1:
'    at atm:
'        use LrHg(1) to calibrate HrHg(0)
'    at LrHg(1)’s full scale:
'        Reg_Pos from 0 to 20; use LrHg(1) to calibrate HrHg(0)
'In case 2:
'    at atm:
'        use LrHg(1) to calibrate HrHg(0)
'    at LrHg(1)’s full scale:
'        Reg_Pos from 0 to 45; use LrHg(1) to calibrate HrHg(0)
'In case 3:
'    at atm:
'        use HrLg(2) to calibrate HrHg(0) and LrHg(1)
'    at HrLg(2)’s full scale:
'        Reg_Pos from 0 to 20; use HrLg(2) to calibrate HrHg(0))
'In case 4:
'    at atm:
'        use HrLg(2) to calibrate HrHg(0) and LrHg(1)
'    at HrLg(2)’s full scale:
'        Reg_Pos from 0 to 20; use LrHg(2) to calibrate LrHg(1)
'    at LrHg(1)’s full scale:
'        Reg_Pos from 20 to 45; use LrHg(1) to calibrate HrHg(0)
'In case 5:
'    at atm:
'        use LrHg(1) to calibrate HrHg(0),HrLg(2),and LrLg(3)
'    at LrLg(3)’s full scale:
'        Reg_Pos from 0 to 3; use LrLg(3) to calibrate HrLg(2)
'    at HrLg(2)’s full scale:
'        Reg_Pos from 3 to 10; use HrLg(2) to calibrate LrHg(1)
'    at LrHg(1)’s full scale:
'        Reg_Pos from 10 to 20; use LrHg(1) to calibrate HrHg(0)
'In case 6:
'    at atm:
'        use LrHg(1) to calibrate HrHg(0),HrLg(2),and LrLg(3)
'    at LrLg(3)’s full scale:
'        Reg_Pos from 0 to 3; use LrLg(3) to calibrate HrLg(2)
'    at HrLg(2)’s full scale:
'        Reg_Pos from 3 to 10; use HrLg(2) to calibrate LrHg(1)
'    at LrHg(1)’s full scale:
'        Reg_Pos from 10 to 20; use LrHg(1) to calibrate HrHg(0)
'In case 7:
'    at atm:
'        use LrHg(1) to calibrate HrHg(0),HrLg(2),and LrLg(3)
'    at LrLg(3)’s full scale:
'        Reg_Pos from 0 to 3; use LrLg(3) to calibrate HrLg(2)
'    at HrLg(2)’s full scale:
'        Reg_Pos from 3 to 10; use HrLg(2) to calibrate LrHg(1)
'    at LrHg(1)’s full scale:
'        Reg_Pos from 10 to 45; use LrHg(1) to calibrate HrHg(0)
'In case 8:
'    at atm:
'        use LrHg(1) to calibrate HrHg(0),HrLg(2),and LrLg(3)
'   at LrLg(3)’s full scale:
'        Reg_Pos from 0 to 3; use LrLg(3) to calibrate HrLg(2)
'    at HrLg(2)’s full scale:
'        Reg_Pos from 3 to 10; use HrLg(2) to calibrate LrHg(1)
'    at LrHg(1)’s full scale:
'        Reg_Pos from 10 to 45; use LrHg(1) to calibrate HrHg(0)
'
' ### necessary changes ###
' ** PGCalibr.frm (PGCalibrate) ** add this form
' ** CAPFLOW.BAS ** changes made on this module:
'    1). add sub CalibratePG
'            sub CheckPG
'            sub Get_Gauge_Case
'            sub TestReadXReturnX4
'            function Gauge_Label$
'    2). change load_Sample:
'        from  If PY1(0) = 0 Then
'                 x_temp% = atm_x4(0) - (P0 / PY2(0) * (PX2(0) - PX1(0)) + PX1(0))
'              If PY1(1) = 0 Then
'                 x_temp% = atm_x4(1) - (P0 / PY2(1) * (PX2(1) - PX1(1)) + PX1(1))
'        to    If PY1(0) < 5 Then
'                 x_temp% = atm_x4(0) - ((P0 - PY1(0)) / (PY2(0) - PY1(0)) * (PX2(0) - PX1(0)) + PX1(0))
'              If PY1(1) < 5 Then
'                 x_temp% = atm_x4(1) - ((P0 - PY1(1)) / (PY2(1) - PY1(1)) * (PX2(1) - PX1(1)) + PX1(1))
' ** CAPWIN.GBL ** add Global varables: Gauge_case%, NN%, GaugeCali_Aborted%
'    when GaugeCali_Aborted = true, quit calibration routine
' ** CAPMAIN.FRM (TitleScrn) **  changes made on this from:
'     1). add a submenu called Calibrate Pressure Gauges under Calibrate menu
'     2). add object PGcali to initiate calibration routine
End Sub

Private Sub testreadxreturnx4(X As Integer)

    If X = 1 And HFLOW% > 1 And vflow% = 0 Then
        vflow% = 1: Move_Valve 9, "O"
    ElseIf X = 1 And HFLOW% < 2 And vflow% = 1 Then
        vflow% = 0: Move_Valve 9, "C"
    End If

    x4 = raw_reading(X)
    If x4 = 25000 And Pres% = 2 And X = 2 Then
       Move_Valve 10, "C": Move_Valve 10, "O"
       x4 = raw_reading(X)
    End If
    If X = 3 Or X > 13 Then Exit Sub
    If X = 0 Then x5 = (x4 - FX1(0, lflow%)) / (FX2(0, lflow%) - FX1(0, lflow%)) * (FY2(0, lflow%) * gasflowconversionfactor - FY1(0, lflow%)) + FY1(0, lflow%)
    
    If X = 1 Then x5 = (x4 - FX1(1, HFLOW%)) / (FX2(1, HFLOW%) - FX1(1, HFLOW%)) * (FY2(1, HFLOW%) * gasflowconversionfactor - FY1(1, HFLOW%)) + FY1(1, HFLOW%)
    
    If X = 2 Then x5 = (x4 - PX1(Pres%)) / (PX2(Pres%) - PX1(Pres%)) * (PY2(Pres%) - PY1(Pres%)) + PY1(Pres%)
    If X = 4 Then
        If Second_Penetrometer And penetrometer_select = 2 Then
            x5 = (x4 - 500) / 20000 * (P2PEN20500 - P2PEN500) + P2PEN500
        Else
            x5 = (x4 - 500) / 20000 * (PEN20500 - PEN500) + PEN500
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    linedemo
End Sub

