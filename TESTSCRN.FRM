VERSION 5.00
Begin VB.Form Testscrn 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autotest"
   ClientHeight    =   9060
   ClientLeft      =   4800
   ClientTop       =   2115
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9060
   ScaleWidth      =   11565
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   11295
      Begin VB.CheckBox chkDivertSampleChamber 
         BackColor       =   &H008080FF&
         Caption         =   "Divert Sample Chamber"
         Height          =   255
         Left            =   5520
         TabIndex        =   61
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton cmdTemperatureControl 
         Caption         =   "Temperature Control"
         Height          =   495
         Left            =   480
         TabIndex        =   60
         Top             =   1230
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdditionalInformation 
         Appearance      =   0  'Flat
         Caption         =   "Additional &Information"
         Height          =   495
         Left            =   8400
         TabIndex        =   59
         Top             =   675
         Width           =   2655
      End
      Begin VB.OptionButton cartridgeOption 
         Caption         =   "Cartridge"
         Height          =   255
         Left            =   2760
         TabIndex        =   58
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton mediaOption 
         Caption         =   "Media"
         Height          =   255
         Left            =   2760
         TabIndex        =   57
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "% Porosity"
         Height          =   255
         Index           =   16
         Left            =   480
         TabIndex        =   56
         Top             =   6240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   210
         Index           =   16
         Left            =   2715
         TabIndex        =   55
         Top             =   6240
         Visible         =   0   'False
         Width           =   8295
      End
      Begin VB.CheckBox tf_check 
         BackColor       =   &H00C0C000&
         Caption         =   "Default (0.715)"
         Height          =   255
         Left            =   8160
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Output File Name"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   2355
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "End User"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   2595
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Test Ref."
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   15
         Top             =   2835
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Sample ID"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   16
         Top             =   3090
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Lot Number"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   17
         Top             =   3330
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Operator"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   18
         Top             =   3570
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Fluid"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   19
         Top             =   3810
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Surface Tension"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   20
         Top             =   4050
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Dry Parameter"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   21
         Top             =   4305
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Wet Parameter"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   22
         Top             =   4545
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Test Type"
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   23
         Top             =   4785
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Diameter"
         Height          =   255
         Index           =   11
         Left            =   480
         TabIndex        =   24
         ToolTipText     =   "For normal samples, this is the diameter of the sample.  For cylindrical samples, this is the outer diameter of the sample holder."
         Top             =   5025
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Cyl. Len. (or 0)"
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   25
         ToolTipText     =   "This value should be 0 for normal samples.  For cylindrical samples, it is the length of the cylinder."
         Top             =   5265
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Thickness"
         Height          =   255
         Index           =   13
         Left            =   480
         TabIndex        =   26
         ToolTipText     =   "For normal samples, this is the thickness of the sample.  For cylindrical samples, this is the thickness of the cylidrical holder."
         Top             =   5520
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Abs. Density"
         Height          =   255
         Index           =   14
         Left            =   480
         TabIndex        =   27
         Top             =   5760
         Width           =   2415
      End
      Begin VB.CheckBox testcap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "Mass"
         Height          =   255
         Index           =   15
         Left            =   480
         TabIndex        =   28
         Top             =   6000
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show &Advanced Settings"
         Height          =   495
         Left            =   5520
         TabIndex        =   6
         Top             =   675
         Width           =   2655
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   15
         Left            =   2715
         TabIndex        =   44
         Top             =   6000
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   14
         Left            =   2715
         TabIndex        =   43
         Top             =   5760
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   13
         Left            =   2715
         TabIndex        =   42
         ToolTipText     =   "For normal samples, this is the thickness of the sample.  For cylindrical samples, this is the thickness of the cylidrical holder."
         Top             =   5520
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   12
         Left            =   2715
         TabIndex        =   41
         ToolTipText     =   "This value should be 0 for normal samples.  For cylindrical samples, it is the length of the cylinder."
         Top             =   5280
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   11
         Left            =   2715
         TabIndex        =   40
         ToolTipText     =   "For normal samples, this is the diameter of the sample.  For cylindrical samples, this is the outer diameter of the sample holder."
         Top             =   5040
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   10
         Left            =   2715
         TabIndex        =   39
         Top             =   4800
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   9
         Left            =   2715
         TabIndex        =   38
         Top             =   4560
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   8
         Left            =   2715
         TabIndex        =   37
         Top             =   4320
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   7
         Left            =   2715
         TabIndex        =   36
         Top             =   4080
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   6
         Left            =   2715
         TabIndex        =   35
         Top             =   3840
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   2715
         TabIndex        =   34
         Top             =   3600
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   2715
         TabIndex        =   33
         Top             =   3360
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   2715
         TabIndex        =   32
         Top             =   3120
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   2715
         TabIndex        =   31
         Top             =   2880
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   2715
         TabIndex        =   30
         Top             =   2640
         Width           =   8295
      End
      Begin VB.CheckBox testval 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   2715
         TabIndex        =   29
         Top             =   2400
         Width           =   8295
      End
      Begin VB.CommandButton Cancelit 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   675
         Width           =   1935
      End
      Begin VB.CommandButton Continue 
         Appearance      =   0  'Flat
         Caption         =   "&Start Test"
         Height          =   495
         Left            =   480
         TabIndex        =   0
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton Help 
         Appearance      =   0  'Flat
         Caption         =   "&Help"
         Height          =   495
         Left            =   8400
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton cmdEditParamFile 
         Appearance      =   0  'Flat
         Caption         =   "&Edit Parameter File"
         Height          =   495
         Left            =   2640
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton cmdPoreShape 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Pore Shape: Circular"
         Height          =   495
         Left            =   5520
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton cmdChgUser 
         Appearance      =   0  'Flat
         Caption         =   "Change &Group"
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   1785
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox tf_value 
         Height          =   285
         Left            =   8520
         TabIndex        =   11
         Text            =   "0.715"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox minp_value 
         Height          =   285
         Left            =   7320
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1635
         Width           =   1095
      End
      Begin VB.TextBox maxp_value 
         Height          =   285
         Left            =   7320
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1995
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   240
         Top             =   6960
      End
      Begin VB.CommandButton cmdEditPreferences 
         Caption         =   "Edit Preferences"
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   675
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Saint Gobain"
         Height          =   495
         Left            =   2640
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diameter and thickness are necessary only for the Pore Count or Envelope Surface Area Analysis"
         ForeColor       =   &H0000FF00&
         Height          =   360
         Index           =   1
         Left            =   240
         TabIndex        =   47
         Top             =   7440
         Visible         =   0   'False
         Width           =   10815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Click (or space bar) on selected value to change value"
         ForeColor       =   &H008080FF&
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   46
         Top             =   7080
         Width           =   9255
      End
      Begin VB.Label tf_label 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         Caption         =   "Tortuosity Factor:"
         Height          =   255
         Left            =   5520
         TabIndex        =   54
         Top             =   795
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Absolute Density and Mass are necessary only for the Envelope Surface Area Analysis"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   48
         Top             =   7800
         Width           =   10815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Pressure:"
         Height          =   255
         Left            =   2880
         TabIndex        =   53
         Top             =   1635
         Width           =   4335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Pressure:"
         Height          =   255
         Left            =   2880
         TabIndex        =   52
         Top             =   1995
         Width           =   4335
      End
      Begin VB.Label p_units_label 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   255
         Left            =   8520
         TabIndex        =   51
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Lines with selected check boxes will be visible in User mode, others will be hidden"
         ForeColor       =   &H008080FF&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   49
         Top             =   8160
         Width           =   10815
      End
      Begin VB.Label Label2 
         Caption         =   "This hidden label is never seen - it is only used as a reference to the width of above boxes"
         Height          =   255
         Left            =   3000
         TabIndex        =   45
         Top             =   6720
         Visible         =   0   'False
         Width           =   7890
      End
   End
End
Attribute VB_Name = "Testscrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim more_to_go As Boolean
Dim toppos_0 As Long
Dim posinc As Long
Dim posoffset As Long
Dim frameoffset As Long
Dim show_advanced As Boolean
Dim dirty As Boolean
Dim original_tf As Single
' seal_state is a flag for determining what text to print for the
' microflow test options line:
' MF_linearSeal = false --> not using linear seal; seal_state = 0, list pressure, time, data time
' MF_linearSeal = true, cyl_len = 0 --> seal_state = 1; list state 0 plus seal diameter
' MF_linearSeal = true, cyl_len > 0 --> seal_state = 2; list state 0 plus inner and outer diameter
Dim seal_state As Integer
'6.71.123.14 Dim ts$(73)                     ' Text strings for this form
Dim ts$(84)                     ' Text strings for this form

Private Sub Cancelit_Click()

    Dim r As VbMsgBoxResult, T As Single
    
    ' pre-figure current tfactor to see if it has changed from original
    If tf_check.value = 0 Then
      T = 0.715
    Else
      T = myVal(tf_value.Text)
    End If
    If T <> original_tf Then dirty = True
    If dirty Then
        r = MsgBox(ts$(1), vbYesNoCancel + vbDefaultButton2, "Capwin")      ' "Do you want to save any changes you made?"
    Else
        r = vbNo
    End If
    If r = vbCancel Then Exit Sub
    RUNNING = False
    If r = vbNo Then
        Unload Me
    Else
        savestuff
    End If
    
End Sub

Private Sub check_micro_seal_type()
' Check to see if user has selected a linear seal for microflow (if microflow is being used).
' If so, set seal_state to indicate which test options to show.

    If MF_linearSeal Then
        If cyl_len(current_unit%) > 0 Then
            seal_state = 2
        Else
            seal_state = 1
        End If
    Else
        seal_state = 0
    End If
    
End Sub

Private Sub set_micro_text()
' Companion to check_micro_seal_type. Once the seal type is set, this routine will
' Rewrite the microflow test options line in order to adjust what choices the user has.
            
    ' First part is standard text that is always shown
    testval(8).Caption = ts$(15) + "=" + Format$(mf_press(current_unit%) * PCNV, "####0.##### ") + PU$ + ", " + ts$(16) + "=" + str$(mf_time(current_unit%)) + " " + ts$(21) + ", " + ts$(20) + " " + ts$(16) + "=" + str$(Step_Time(current_unit%)) + " " + ts$(21)    ' "Pressure"/"Time"/"min"/"Data"/"Time"/"min"
            
    If seal_state = 0 Then Exit Sub
    
    If seal_state = 1 Then
        testval(8).Caption = testval(8).Caption + ", " + ts$(70) + "=" + str$(MF_sealDiam) + " cm"          ' "Seal Diameter"
    ElseIf seal_state = 2 Then
        testval(8).Caption = testval(8).Caption + ", " + "ID" + "=" + str$(MF_innerDiam) + " cm" + ", " + "OD" + "=" + str$(MF_outerDiam) + " cm"
    End If


End Sub
Private Sub setup_testvals(u$)

    Dim i As Integer
    
    load_user_stuff u$
    checkLegalSettings
    update_sample_ID
    update_filename
    If tfactor <= 0 Then tfactor = 0.715
    If tfactor = 0.715 Then
      tf_check = 0
    Else
      tf_check = 1
      tf_value.Text = str$(tfactor)
      ' we need to hide this because setting tf_check
      ' to 1 will make this visible and we don't want
      ' it to be visible until they press the advanced
      ' button
      tf_value.Visible = False
    End If
    original_tf = tfactor
    minp_value.Text = Format$(minp_set(current_unit%) * PCNV, "######0.####")
    maxp_value.Text = Format$(maxp_set(current_unit%) * PCNV, "######0.####")
    If minmaxunits = "p" Then         ' using pressure
        p_units_label.Caption = ts$(2) + " " + PU$          ' "Pressure in"
    ElseIf TType%(current_unit%) < 3 Then                   ' using pore diameter, but only for CFP or bubble point
        ' 6.71.61 rvw - don't do pressure unit conversion when displaying micron size - default pressure
        '  unit of PSI is correct for internal values.  This was causing a problem in Europe or anywhere where they
        '  set the pressure units to anything other than PSI.
        '  Also, if minimum pressure is set to 0, set microns to 2000.
        '  Also, adjust format so number displays better
        p_units_label.Caption = "microns"
        If minp_set(current_unit%) = 0 Then
            maxp_value.Text = "2000"
        Else
            maxp_value.Text = Format$((0.415 * surfTen(current_unit%)) / minp_set(current_unit%), "###0.####")
        End If
        minp_value.Text = Format$((0.415 * surfTen(current_unit%)) / maxp_set(current_unit%), "###0.####")
    End If
    If TMode%(current_unit%) < 0 Then
      TMode%(current_unit%) = -TMode%(current_unit%)
      diffpgflow(current_unit%) = True
      If TType%(current_unit%) = 3 Then
        ' gas perm with negative tmode is actually microflow gas perm with tmode=0
        TMode%(current_unit%) = 0
      End If
    Else
      diffpgflow(current_unit%) = False
    End If
    If GasPerm And TType%(current_unit%) < 3 Then
      TType%(current_unit%) = 3
      If TMode%(current_unit%) = 1 Or TMode%(current_unit%) > 2 Then TMode%(current_unit%) = 0
    End If
    If WESA_exclusive Then
        TType%(current_unit%) = 3
        TMode%(current_unit%) = 1
    End If
    
    If Not WESA_enabled Then
        ' JF 12-30-09
        ' Changed TMode%(current_unit%) <> 2  to  TMode%(current_unit%) <> 1
        If TType%(current_unit%) <> 3 And TMode%(current_unit%) <> 1 Then
            testval(14).Enabled = False
            testval(15).Enabled = False
        End If
        If TType%(current_unit%) = 3 And TMode%(current_unit%) = 1 Then TMode%(current_unit%) = 0
    End If
    
    If itester Then
      Rem check for types of tests that can't be done
      If TType%(current_unit%) = 1 Or TType%(current_unit%) = 3 Then
        TType%(current_unit%) = 2
        TMode%(current_unit%) = 1
      End If
    End If
    If BPTester Then
        ' 6.71.38x begin
        If hydrohead_exclusive Then
            ' for hydrohead exclusive machine, test is always hydrohead
            TType%(current_unit%) = 2
            TMode%(current_unit%) = 3
        ElseIf TType%(current_unit%) = 5 Then
            TMode%(current_unit%) = 1
        ElseIf TType%(current_unit%) <> 2 Then
            ' otherwise, if they don't have a bubblepoint-type test selected
            ' force it to normal bubble point
            TType%(current_unit%) = 2
            TMode%(current_unit%) = 1
        ElseIf TMode%(current_unit%) > 1 Then
            ' otherwise, if they have a bubblepoint-type test selected but it is
            ' not mode 0 or 1 (or 3 if hydrohead is enabled) set it back to 1
            If TMode%(current_unit%) <> 3 Or hydrohead = False Then
                TMode%(current_unit%) = 1
            End If
        ' 6.71.38x end
        End If
    End If
    If (Not fluidsensor) And TType%(current_unit%) = 1 Then
     If TMode%(current_unit%) = 6 Or TMode%(current_unit%) = 7 Then
      TMode%(current_unit%) = TMode%(current_unit%) - 4
     ElseIf TMode%(current_unit%) = 9 Then
      TMode%(current_unit%) = 8
     End If
    End If
    For i = 0 To 15
        testcap(i).value = qcshow(i)
    Next i
    'qcshow 16 through 20 were taken by other things '6.71.123.14
    
    'For i = 16 To 21 '6.71.123.14
    '    testcap(i).value = qcshow(i + 5)
    'Next i
    
    If Not SqrPore Then
        cmdPoreShape.Visible = False
        SquarePores = False
    End If

    ' test values can not be more than 85 characters
    testval(0).Caption = make_sure(shorten(OutFilename$(current_unit%)))
    testval(1).Caption = make_sure(Line1$(current_unit%))
    testval(2).Caption = make_sure(Line2$(current_unit%))
    testval(3).Caption = make_sure(sid$(current_unit%))
    testval(4).Caption = make_sure(lot_number$(current_unit%))
    testval(5).Caption = make_sure(operator$(current_unit%))
    testval(7).Caption = str$(surfTen(current_unit%)) + " " + ts$(44)       ' "dynes/cm"
    ' number 8 can sometimes be a different caption
    testcap(8).Caption = "    " + ts$(3)            ' "Dry Parameter"
    testval(8).Caption = make_sure(shorten(TPFDRY$(current_unit%)))
    testval(9).Caption = make_sure(shorten(TPFWET$(current_unit%)))
    If TType%(current_unit%) = 1 Then
        If CF_SampleType = 0 Then
            testval(11).Caption = Format$(Diam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
            testval(12).Caption = Format$(cyl_len(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
            testval(13).Caption = Format$(thick(current_unit%) / thick_unit_conversion#, "####0.##### ") + thick_unit_name$
        Else
            testval(11).Caption = Format$(innerDiam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
            testval(12).Caption = Format$(outerDiam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
            testval(13).Caption = Format$(cyl_len(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
        End If
    Else
        testval(11).Caption = Format$(Diam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
        testval(12).Caption = Format$(cyl_len(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
        testval(13).Caption = Format$(thick(current_unit%) / thick_unit_conversion#, "####0.##### ") + thick_unit_name$
    End If
    testval(14).Caption = Format$(density(current_unit%) / dens_unit_c, "####0.#####") + dens_unit
    If SCDiam > Diam(current_unit%) Then
        testval(15).Caption = Format$(mass(current_unit%) / mass_unit_c, "####0.#####") + mass_unit + ";  " + ts$(4) + ": " + Format$(SCDiam / linear_unit_conversion#, "####0.##### ") + linear_unit_name$ ' "Sample Cut Diameter for Measured Mass"
    Else
        testval(15).Caption = Format$(mass(current_unit%) / mass_unit_c, "####0.#####") + mass_unit
    End If
    testval(10).Caption = TestName$()
    testval(16).Caption = Format$(TargetPercPorosity, "####0.#####")                '6.71.123.14
    
    'If slurry_tube_exists Then  'AJB 09-21-09
    '    testval(17).Caption = Format$(BuildCakeByPressureTargetPressure, "####0.#####") '6.71.123.14
    '    testval(18).Caption = Format$(BuildCakeByFlowTargetFlow, "####0.#")             '6.71.123.14
    '    testval(19).Caption = Format$(BuildCakeByFlowEndPressure, "####0.#####")        '6.71.123.14
    '    testval(20).Caption = Format$(SlurryTubeWashCycleTargetFlow, "####0.#")         '6.71.123.14
    '    testval(21).Caption = Format$(BuildCakeByPressureEndFlow, "####0.####")
    'End If
    setup_testname
    
End Sub
Private Sub setup_testname()
' note that this routine uses the enabled value of testval (and label1)
' to determine if that line should show on the screen or not.  The
' mungscreen routine uses the enabled value when setting up the screen.
' We have to do this because sometimes this routine is called when
' the form hasn't been fully drawn yet, and the visible parameter of
' everything is set to false temporarily.

Dim show_min_max_p As Boolean   ' true if the type of test uses the min and max pressure values
Dim show_tfactor As Boolean     ' true if the type of test uses the tortuosity factor
Dim liqmid As String
Dim liqleft As String


testcap(8).Caption = "    " + ts$(3)        ' "Dry Parameter"
testval(8).Caption = make_sure(shorten(TPFDRY$(current_unit%)))
show_min_max_p = True
show_tfactor = True

Select Case TType%(current_unit%)
    Case 1
        testcap(6).Caption = " " + ts$(5)   ' "Fluid"
        testval(6).Caption = make_sure(fluid$(current_unit%))
        testval(7).Enabled = True
        If TMode%(current_unit%) = 4 Or TMode%(current_unit%) = 10 Or TMode%(current_unit%) = 11 Then
            testval(8).Enabled = False
        Else
            testval(8).Enabled = True
            If Dir$(TPFDRY$(current_unit%)) = "" Then
                MsgBox ts$(7) + " " + ts$(3) + " " + ts$(8)       ' "Warning: The dry parameter file is invalid"
            End If
        End If
        testval(9).Enabled = True
        If Dir$(TPFWET$(current_unit%)) = "" Then
            MsgBox ts$(7) + " " + ts$(6) + " " + ts$(8)             ' "Warning: The wet parameter file is invalid"
        End If
        cmdEditParamFile.Enabled = True
        If CF_SampleType = 0 Then
            testcap(11).Caption = "Diameter"
            testval(11).Enabled = True
            testval(12).Enabled = False
            testcap(13).Caption = "Thickness"
            testval(13).Enabled = Not piston_position_transducer_exists
        Else
            testcap(11).Caption = "Inner Diameter"
            testcap(12).Caption = "Outer Diameter"
            testcap(13).Caption = "Sample Height"
            testval(11).Enabled = True
            testval(12).Enabled = True
            testval(13).Enabled = True
        End If
        testval(14).Enabled = False
        testval(15).Enabled = False
        
        Label1(1).Enabled = True
        Label1(2).Enabled = True
        testval(16).Enabled = piston_position_transducer_exists '% Porosity '6.71.123.14
    Case 2
        testcap(6).Caption = " " + ts$(5)           ' "Fluid"
        testval(6).Caption = make_sure(fluid$(current_unit%))
        testval(7).Enabled = True
        testval(8).Enabled = False
        testval(9).Enabled = True
        If Dir$(TPFWET$(current_unit%)) = "" Then
            MsgBox ts$(7) + " " + ts$(6) + " " + ts$(8)     ' "Warning: The wet parameter file is invalid"
        End If
        cmdEditParamFile.Enabled = True
        testval(11).Enabled = False
        testval(12).Enabled = False
        testval(13).Enabled = False
        Label1(1).Enabled = False
        testval(14).Enabled = False
        testval(15).Enabled = False
        Label1(2).Enabled = False
    Case 3
        show_tfactor = False
        testcap(6).Caption = "   " + ts$(9)     ' "Gas"
        If Len(Gas$) < 21 Then
            testval(6).Caption = make_sure(Gas$)
        Else
            testval(6).Caption = make_sure(Mid$(Gas$, 20) + " mu=" + str$(val(Left$(Gas$, 19))))
        End If
        testval(7).Enabled = False
        If TMode%(current_unit%) = 0 Then
            ' if we are doing microflow gas permeability, don't change testval(10).caption
            ' because it is already set properly
            If diffpgflow(current_unit%) = False Then
                testval(10).Caption = ts$(10)       ' "Gas Permeability"
            End If
        ElseIf TMode%(current_unit%) = 1 Then
            testval(10).Caption = ts$(11)       ' "Envelope Surface Area Analysis"
        ElseIf TMode%(current_unit%) = 2 Then
            testval(10).Caption = ts$(12)       ' "Pressure Drop Test"
        ElseIf TMode%(current_unit%) = 10 Then
            testval(10).Caption = ts$(80)
        ElseIf TMode%(current_unit%) = 11 Then
            testval(10).Caption = ts$(81)
        ElseIf TMode%(current_unit%) = 12 Then
            testval(10).Caption = ts$(82)
        ElseIf TMode%(current_unit%) = 20 Then
            testval(10).Caption = ts$(84)
        End If
        testval(8).Enabled = True
        If Dir$(TPFDRY$(current_unit%)) = "" Then
            MsgBox ts$(7) + " " + ts$(3) + " " + ts$(8)     ' "Warning: The dry parameter file is invalid"
        End If
        testval(9).Enabled = False
        cmdEditParamFile.Enabled = True
        testval(11).Enabled = True
        testval(12).Enabled = True
        '6.71.123.14 testval(13).Enabled = True
        testval(13).Enabled = Not piston_position_transducer_exists 'Thickness '6.71.123.14
        Label1(1).Enabled = False
        Label1(2).Enabled = False
        If TMode%(current_unit%) <> 0 Then
            testval(14).Enabled = True
            testval(15).Enabled = True
        Else
            testval(14).Enabled = False
            testval(15).Enabled = False
        End If
        testval(16).Enabled = piston_position_transducer_exists '% Porosity '6.71.123.14
    Case 4
        show_tfactor = False
        show_min_max_p = False
        testcap(6).Caption = ts$(13)        ' "Liquid"
        
        If Len(Liquid$(current_unit%)) < 21 Then
            testval(6).Caption = make_sure(Liquid$(current_unit%))
        Else
            testval(6).Caption = make_sure(Mid$(Liquid$(current_unit%), 20) + " mu=" + str$(val(Left$(Liquid$(current_unit%), 19))))
        End If
        testval(7).Enabled = False
        testval(8).Enabled = False
'        If TMode%(current_unit%) = 1 And special_ambient Then
'            testval(9).Enabled = True
'            cmdEditParamFile.Enabled = True
'            If Dir$(TPFWET$(current_unit%)) = "" Then
'                MsgBox ts$(7) + " " + ts$(6) + " " + ts$(8)         ' "Warning: The wet parameter file is invalid"
'            End If
'        Else
            testval(9).Enabled = False
            cmdEditParamFile.Enabled = True
'        End If
        testval(11).Enabled = True
        testval(12).Enabled = True
        '6.71.123.14 testval(13).Enabled = True
        testval(13).Enabled = Not piston_position_transducer_exists 'Thickness '6.71.123.14
        Label1(1).Enabled = False
        testval(14).Enabled = False
        testval(15).Enabled = False
        Label1(2).Enabled = False
    Case 5
        show_tfactor = False
        show_min_max_p = False
        testcap(6).Caption = "   " + ts$(9)             ' "Gas"
        If Len(Gas$) < 21 Then
            testval(6).Caption = make_sure(Gas$)
        Else
            testval(6).Caption = make_sure(Mid$(Gas$, 20) + " mu=" + str$(val(Left$(Gas$, 19))))
        End If
        testval(7).Enabled = False
        testcap(8).Caption = ts$(14)                    ' "Hold Parameters"
        testval(8).Caption = ts$(15) + "=" + Format$(Hold_Press(current_unit%) * PCNV, "####0.##### ") + PU$ + ", " + ts$(16) + "=" + str$(Hold_Time(current_unit%)) + " " + ts$(17) + ", " + ts$(18) + "=" + str$(Hold_Delay(current_unit%)) + " sec."     ' "Pressure"/"Time"/"sec"/"Delay"
        If Hold_Rate(current_unit%) > 0 Then
            testval(8).Caption = testval(8).Caption + ", " + ts$(19) + "=" + Format$(Hold_Rate(current_unit%) * PCNV, "###0.##### ") + PU$ + IIf(PH_fail_method$ = "dpdt", "/" + IIf(pressHoldUnit = "S", "sec", "min"), "") ' "Fail"/"sec"
        End If
        testval(8).Enabled = True
        testval(9).Enabled = False
        cmdEditParamFile.Enabled = False
        testval(11).Enabled = False
        testval(12).Enabled = False
        testval(13).Enabled = False
        Label1(1).Enabled = False
        testval(14).Enabled = False
        testval(15).Enabled = False
        Label1(2).Enabled = False
    Case 6
        show_tfactor = False
        show_min_max_p = False
        If lvperm_enable Then
            testcap(6).Caption = ts$(13)        ' "Liquid"
            testval(6).Caption = make_sure(fluid$(current_unit%))
        Else
            testcap(6).Caption = "   " + ts$(9) ' "Gas"
            If Len(Gas$) < 21 Then
                testval(6).Caption = make_sure(Gas$)
            Else
                testval(6).Caption = make_sure(Mid$(Gas$, 21) + " mu=" + str$(val(Left$(Gas$, 20))))
            End If
        End If
        testval(7).Enabled = False
        testcap(8).Caption = ts$(14)        ' "Hold Parameters"
        If lvperm_enable Then
            testval(8).Caption = ts$(16) + "=" + str$(mf_time(current_unit%)) + " " + ts$(21) + ", " + ts$(20) + " " + ts$(16) + "=" + str$(Step_Time(current_unit%)) + " " + ts$(21) ' "Time"/"min"/"Data"/"Time"/"min"
        Else
            ' Regular microflow test
            check_micro_seal_type
            set_micro_text
        End If
        testval(8).Enabled = True
        testval(9).Enabled = False
        cmdEditParamFile.Enabled = False
        testval(11).Enabled = True
        testval(12).Enabled = True
        '6.71.123.14 testval(13).Enabled = True
        testval(13).Enabled = Not piston_position_transducer_exists 'Thickness '6.71.123.14
        Label1(1).Enabled = False
        testval(14).Enabled = False
        testval(15).Enabled = False
        Label1(2).Enabled = False
    Case 7
        show_tfactor = False
        show_min_max_p = False
        testcap(6).Caption = ts$(13)        ' "Liquid"
        
        If Len(Liquid$(current_unit%)) < 21 Then
            testval(6).Caption = make_sure(Liquid$(current_unit%))
        Else
            testval(6).Caption = make_sure(Mid$(Liquid$(current_unit%), 20) + " mu=" + str$(val(Left$(Liquid$(current_unit%), 19))))
        End If
        testval(7).Enabled = False
        testval(8).Enabled = False
'        If TMode%(current_unit%) = 1 And special_ambient Then
'            testval(9).Enabled = True
'            cmdEditParamFile.Enabled = True
'            If Dir$(TPFWET$(current_unit%)) = "" Then
'                MsgBox ts$(7) + " " + ts$(6) + " " + ts$(8)         ' "Warning: The wet parameter file is invalid"
'            End If
'        Else
            testval(9).Enabled = False
            cmdEditParamFile.Enabled = True
'        End If
        testval(11).Enabled = True
        testval(12).Enabled = True
        '6.71.123.14 testval(13).Enabled = True
        testval(13).Enabled = Not piston_position_transducer_exists 'Thickness '6.71.123.14
        Label1(1).Enabled = False
        testval(14).Enabled = False
        testval(15).Enabled = False
        Label1(2).Enabled = False
End Select
minp_value.Visible = show_min_max_p
maxp_value.Visible = show_min_max_p
Label4.Visible = show_min_max_p
Label5.Visible = show_min_max_p
p_units_label.Visible = show_min_max_p
tf_check.Enabled = show_tfactor ' mungscreen takes care of this

mungscreen

End Sub

Private Sub cartridgeOption_Click()
    mediaOption.value = False
    cartridgeOption.value = True
    current_unit% = 2
    cartridge_tester_side = 1
    'load_user_unit_stuff 2
    'setup_testvals 2
    
End Sub

Private Sub cartridgeOption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mediaOption.value = False
    cartridgeOption = True
    setup_testvals 2
End Sub

Private Sub cmdAdditionalInformation_Click()
    If useAdditionalInfo And numberOfAdditionalInfoLines > 0 Then
        frmAdditionalInformation.Show 1
    Else
        MsgBox "Additional Information has not been setup.  Please setup and then click the Additional Information button again.", vbInformation
        prefsForm.SetCurrentTab 11
        prefsForm.Show 1
    End If
End Sub

Private Sub cmdChgUser_Click()

Dim DFile$, T%, XTemp$
Dim u$

If current_unit% = 1 Then
    u$ = ""
Else
    u$ = Format$(current_unit%)
End If
    T_Select$ = "USER"
    Selection.Show 1
    If T_Select$ = "Cancel" Then Exit Sub
    WPPS "default", "user", Curr_U$, IFile$
    setup_testvals u$

End Sub

Private Sub cmdEditParamFile_Click()
    If TType(current_unit%) = 7 Then
        resin_parameters.Show 1
    ElseIf TType%(current_unit%) = 4 Then
        autoparm.Show 1
    Else
        TestScreenVisible = True
        autoparm.Show 1
        TestScreenVisible = False
    End If
End Sub

Private Sub cmdEditPreferences_Click()
    TestScreenVisible = True
    prefsForm.Show 1
    TestScreenVisible = False
    ' just in case, check the seal length variable as they may have changed it
    If TType%(current_unit%) = 6 And lvperm_enable = False Then
        check_micro_seal_type
        set_micro_text
    End If
End Sub

Private Sub cmdPoreShape_Click()

    If cmdPoreShape.Caption = ts$(22) Then              ' "&Pore Shape: Circular"
        cmdPoreShape.Caption = ts$(23)                  ' "&Pore Shape: Square"
        SquarePores = True
    Else
        cmdPoreShape.Caption = ts$(22)                  ' "&Pore Shape: Circular"
        SquarePores = False
    End If

End Sub

Private Sub cmdTemperatureControl_Click()
    frmTemperatureControl.Show 1
End Sub

Private Sub Command1_Click()
' turn on advanced settings mode
    show_advanced = True
    mungscreen
End Sub

Private Sub mungscreen()

    Dim L As Long, i As Integer
    Dim test1 As Integer
    Dim test2 As Boolean

    ' command1 (the advanced button) only shows up if you are not advanced
    ' and you are in supervisor mode
    Command1.Visible = (Not show_advanced) And supervisor
    ' the edit parameter file button only shows up in advanced mode
    cmdEditParamFile.Visible = show_advanced
    cmdEditPreferences.Visible = show_advanced
    ' the help button only shows up in advanced mode
    Help.Visible = show_advanced
    ' the SqrPore button only shows up if it is enabled and you are in advanced mode
    cmdPoreShape.Visible = SqrPore And show_advanced
    ' the user button only shows up in advanced mode
    cmdChgUser.Visible = show_advanced
    ' the tortuosity factor stuff only shows up in advanced mode
    ' and then only if the check box is enabled
    tf_label.Visible = show_advanced And tf_check.Enabled
    tf_check.Visible = show_advanced And tf_check.Enabled
    tf_value.Visible = (tf_check.value = 1) And show_advanced And tf_check.Enabled
    
    If show_advanced Then
        Frame1.BackColor = &HC0C000
    Else
        Frame1.BackColor = &H8080FF
    End If
    
    ' now start putting the testcap and testval elements on
    L = toppos_0
    '6.71.123.14 For i = 0 To 15
    For i = 0 To 15
        test2 = testval(i).Enabled
        test1 = testcap(i).value
        If testval(i).Enabled And (show_advanced Or testcap(i).value = 1) Then
            testcap(i).top = L
            testval(i).top = L + posoffset
            L = L + posinc
            ' change color in advanced mode
            If show_advanced Then
                testcap(i).BackColor = &HC0C000
            End If
            testcap(i).Enabled = True
            testcap(i).Visible = True
            testval(i).Visible = True
            ' so that we can always see that the value has focus, we
            ' always make sure that it has at least one character in it
            If Len(testval(i).Caption) = 0 Then
                testval(i).Caption = " "
            End If
        Else
    '        testcap(i).Top = toppos_0 + posinc * 20 ' should be far enough down
    '        testval(i).Top = testcap(i).Top ' offset doesn't matter - it can't be seen
            testcap(i).Visible = False
            testval(i).Visible = False
        End If
    Next i
    
    ' label about check boxes are only in advanced mode
    Label1(3).Enabled = show_advanced
    ' labels start at 1/2 increment down
    L = L + posinc / 2
    For i = 0 To 3
        If Label1(i).Enabled Then
            Label1(i).top = L
            L = L + posinc * 2
            Label1(i).Visible = True
        Else
            Label1(i).Visible = False
        End If
    Next i

    ' end of screen is 2 increment down
    L = L + posinc * 2                      ' *2 added when items were widened and rearranged for localization
    Me.Height = L + posinc
    Frame1.Height = Me.Height - frameoffset
    Me.top = (Screen.Height - Height) / 2
    If Me.top < 0 Then Me.top = 0

End Sub

Private Sub Command2_Click()

    freePressure.Show 1
    
End Sub

Private Sub Continue_Click()
    RUNNING = True
    savestuff
End Sub

Private Sub savestuff()

    Dim i As Integer, u$
    
    Rem check hydrohead test
    If TType%(current_unit%) = 2 And TMode%(current_unit%) = 3 And (te_number% And 2) = 0 Then
        MsgBox ts$(24)          ' "Your instrument does not have hydrohead test capabilities."
        Exit Sub
    End If
    Rem check parameter files
'    If TType%(current_unit%) = 1 Or TType%(current_unit%) = 2 Or (TType%(current_unit%) = 4 And TMode%(current_unit%) = 1 And special_ambient) Then
    If TType%(current_unit%) = 1 Or TType%(current_unit%) = 2 Then
        If Dir$(TPFWET$(current_unit%)) = "" Then
            MsgBox ts$(6) + " " + ts$(8) + " " + ts$(25) ' "The wet parameter file is invalid - please check"
            Exit Sub
        End If
    End If
    If TType%(current_unit%) = 3 Or (TType%(current_unit%) = 1 And TMode%(current_unit%) <> 4 And TMode%(current_unit%) <> 10 And TMode%(current_unit%) <> 11) Then
        If Dir$(TPFDRY$(current_unit%)) = "" Then
            MsgBox ts$(3) + " " + ts$(8) + " " + ts$(25)    ' "The dry parameter file is invalid - please check"
            Exit Sub
        End If
    End If
    Rem check wesa test
    If TType%(current_unit%) = 3 And TMode%(current_unit%) = 1 And (te_number% And 1) = 0 Then
        MsgBox ts$(26)                                  ' "Your instrument does not have envelope surface area test capabilities."
        Exit Sub
    End If
    'check pressure drop test
    If TType%(current_unit%) = 3 And TMode%(current_unit%) = 2 And (te_number% And 4) = 0 Then
        MsgBox ts$(27)                                  ' "Your instrument does not have pressure drop test capabilities."
        Exit Sub
    End If
    
    If RUNNING And Dir$(OutFilename$(current_unit%)) <> "" Then
        If LCase$(Right$(OutFilename$(current_unit%), 25)) <> "do not save data file.cft" And simpleqc_enable = False Then
            reply% = MsgBox(ts$(28) + vbCrLf + OutFilename$(current_unit%), 308, "Test Setup")      ' "The selected File Already exists. Overwrite?"
            If reply% = 7 Then Exit Sub
        End If
    End If

    If myVal(minp_value.Text) <= 0 Then minp_value.Text = "0"
    If myVal(maxp_value.Text) <= 0 Then maxp_value.Text = "0"
    
    If myVal(minp_value.Text) >= myVal(maxp_value.Text) Then
        MsgBox ts$(30)                      ' "Maximum pressure must be greater than minimum pressure"
        Exit Sub
    End If
    
    If minmaxunits = "p" Or TType%(current_unit%) >= 3 Then
        minp_set(current_unit%) = myVal(minp_value.Text) / PCNV
        maxp_set(current_unit%) = myVal(maxp_value.Text) / PCNV
    ElseIf minmaxunits = "d" And TType%(current_unit%) < 3 Then       ' need to convert pore diameter into pressure
        ' 6.71.61 rvw - remove pressure unit conversion for diameter - needs to remain in internal units of
        '  PSI no matter what the user units are set to.
        ' Note: the receiving variables are deliberately reversed: min pore diameter will give max pressure,
        ' and vice versa
        If myVal(minp_value.Text) = 0 Then
            maxp_set(current_unit%) = Maxpres
        Else
            maxp_set(current_unit%) = (0.415 * surfTen(current_unit%)) / (myVal(minp_value.Text))
        End If
        ' Anything over 2000 microns sets to 0 PSI starting pressure to avoid situation where
        ' the program tries to do a starting pressure of some really small pressure when it should
        ' actually start at 0.
        If myVal(maxp_value.Text) >= 2000 Then
            minp_set(current_unit%) = 0
        Else
            minp_set(current_unit%) = (0.415 * surfTen(current_unit%)) / (myVal(maxp_value.Text))
        End If
    End If

    ' 6.71.22a begin
    If autocompress And Compression_Increase_Factor < 0 And compression_pressure <> 0 And maxp_set(current_unit%) * Abs(Compression_Increase_Factor) > compression_pressure Then
        MsgBox ts$(73)      ' "Maximum pressure too high for safe operation using current compression pressure"
        Exit Sub
    End If
    ' 6.71.22a end
    
    If tf_check.value = 0 Then
      tfactor = 0.715
    Else
      tfactor = myVal(tf_value.Text)
    End If
    
    'If numberOfMicrometers > 0 Then
    '    readMicrometer currentMicrometer
    '    Diam(current_unit%) = micrometers(currentMicrometer).dblValue
    'End If
    
    ' save the current user stuff that they just setup
    If current_unit% = 1 Then u$ = "" Else u$ = Format$(current_unit%)
    If diffpgflow(current_unit%) Then
        ' negating tmode is not for gas perm diffpgflow
        If TType%(current_unit%) <> 3 Then
            TMode%(current_unit%) = -TMode%(current_unit%)
        Else
            ' for gas perm diffpgflow, tmode must be 0
            ' and the resulting tmode to save to the user file is -1
            If TMode%(current_unit%) = 0 Then
                TMode%(current_unit%) = -1
            End If
        End If
    End If
    
    If TType%(current_unit%) = 1 And CF_SampleType = 1 Then
        Diam(current_unit%) = innerDiam(current_unit%)
        thick(current_unit%) = (outerDiam(current_unit%) - innerDiam(current_unit%)) / 2
    End If
    
    save_user_stuff u$
    ' if chamber 1 is not selected, we had better save the global stuff as well
    If Not selchamber(1) Then
        save_user_global_stuff
    End If
    ' now that we are done saving the tmode to the user stuff, we want to set it back
    ' to what it had been.  If it was a gas perm with a -1 then we don't want to
    ' set it to 1, but we want to set it back to 0
    If TType%(current_unit%) = 3 And TMode%(current_unit%) = -1 Then
        TMode%(current_unit%) = 0
    End If
    ' for other cases of a negative tmode, we just want the absolute value
    TMode%(current_unit%) = Abs(TMode%(current_unit%))

    If more_to_go Then
        Do
            current_unit% = current_unit% + 1
            If selchamber(current_unit%) Then
                Form_Load
                Exit Sub
            End If
        Loop Until current_unit% >= chambers
        ' if we fall through here then something is wrong
        MsgBox ts$(31)              ' "Warning: I thought there were more chambers to setup, but I can't find any.  Please let PMI know that you have seen this message"
    End If
    TitleScrn.MousePointer = 11
    Aborted = False ' set to signify that we are going to run the test
    ' set current_unit% to the first selected unit for testing
    current_unit% = 1
    nowait_at_beginning = False
    nowait_at_end = False
    save_setup_data_flag = True
    
    If cartridge_tester Then
        'ajb 10-30-09
        If mediaOption.value = True Then
            WPPS "Capstuff", "cartridge_tester_side", "0", CSFile$
            current_unit% = 1
            cartridge_tester_side = 0
        Else
            WPPS "Capstuff", "cartridge_tester_side", "1", CSFile$
            current_unit% = 2
            cartridge_tester_side = 1
        End If
    
        'ajb 11-01-09
        If TType%(current_unit%) = 4 And cartridgeOption.value = True Then
            penetrometer_select = 2
        End If
        process_CVtable
    End If
    
    If sampleChamberDiverterValve >= 0 Then
        divertSampleChamber = IIf(chkDivertSampleChamber.value = 0, False, True)
        If divertSampleChamber Then
            WPPS "Capstuff", "DivertSampleChamber", "1", CSFile$
        Else
            WPPS "Capstuff", "DivertSampleChamber", "0", CSFile$
        End If
    End If
    
'    If testval(6).Caption = "Nitrogen" Then
'    Gas$ = "0.01770029092057747" + Space$(20) + "Nitrogen"
'    End If
'    If testval(6).Caption = "Water" Then
'    fluid$(current_unit%) = "0.89" + Space$(20) + "Water"
'    End If
'    If testval(6).Caption = "Air" Then
'    Gas$ = "0.018615979546772812" + Space$(20) + "Air"
'    End If
'    If testval(6).Caption = "Galwick" Then
'    Liquid$(current_unit%) = "8.0" + Space$(20) + "Galwick"
'    End If
    Unload Me
    
End Sub

Private Sub Form_Load()

Dim u$, i As Integer

LoadTextStrings

' Superseded by LoadTextStrings:
'For i = 0 To 15
'    testval(i).FontName = system_font
'    testval(i).fontsize = font_size
'    testval(i).fontbold = font_bold
'Next i

If simpleqc_enable Then
    Me.WindowState = 0 ' surik changed 1 to 0
Else
    Me.WindowState = 0
End If
If current_unit = 1 Then
    u$ = ""
Else
    u$ = Format$(current_unit)
End If
If first_test_setup Then
    ' things that get initialized only for the first load of the form
    ' initialize some positioning information:
    toppos_0 = testcap(0).top
    posinc = testcap(1).top - testcap(0).top
    posoffset = testval(0).top - testcap(0).top
    frameoffset = Me.Height - Frame1.Height
    show_advanced = (auto_advanced And supervisor)
    dirty = False
    first_test_setup = False
End If
more_to_go = False
If multiChamberSystem And manualMultiChamber = False Then
    Me.Caption = ts$(32) + " " + str$(current_unit)         ' "Auto Test Setup: Chamber"
    Rem check to see if there are more units to set up
    Me.Continue.Caption = ts$(33)                           ' "&Start Test"
    If current_unit < chambers Then
        For i = current_unit + 1 To chambers
            If selchamber(i) Then
                Me.Continue.Caption = ts$(34)               ' "&Next Unit"
                more_to_go = True
                Exit For
            End If
        Next i
    End If
End If

    ' these don't do anything any more
    'RemoveSysMenu Me
    'TitleScrn.Enabled = False
    Dim temp As String * 1
    Me.Left = (Screen.Width - Width) / 2
    If Me.Left < 0 Then Me.Left = 0
    RUNNING = False
    
    setup_testvals u$
    
    TitleScrn.MousePointer = 0
    
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(35) + str$(unitnumber)      ' "Unit"
    End If
If simpleqc_enable Then
    Timer1.Enabled = True
End If

' Select pore diameter or pressure for endpoints; default is pressure
If minmaxunits = "d" And TType%(current_unit%) < 3 Then             ' Use pore diameter, but only for test types with surf. ten.
    Label4.Caption = ts$(68)
    Label5.Caption = ts$(69)
End If

'AJB 10-30-09
If cartridge_tester Then
    mediaOption.Visible = True
    cartridgeOption.Visible = True
    If cartridge_tester_side = 0 Then
        mediaOption.value = True
        cartridgeOption.value = False
        current_unit% = 1
    Else
        mediaOption.value = False
        cartridgeOption.value = True
        current_unit% = 2
    End If
End If

If sampleChamberDiverterValve >= 0 Then
    chkDivertSampleChamber.Visible = True
    chkDivertSampleChamber.BackColor = Frame1.BackColor
    chkDivertSampleChamber.value = IIf(divertSampleChamber, 1, 0)
Else
    chkDivertSampleChamber.Visible = False
End If

'Add ability to hide temperature control button
    If useTemperatureControlForAuto(current_unit%) Then
        cmdTemperatureControl.Visible = True
    Else
        cmdTemperatureControl.Visible = False
    End If

' Added 6.71.57
dirty = False
'edc 12-11-06 alter border color and caption
Me.Caption = Me.Caption & "    " & SubCaption
Me.BackColor = lngBorderColor

End Sub

Private Sub Form_Unload(cancel As Integer)
    ' this doesn't do anything
    'TitleScrn.Enabled = True
End Sub

Private Sub Help_Click()
    
    Dim r As Long
    
    r = WinHelp(hwnd, HelpFile$, Help_Context, ByVal 45&)

End Sub

Private Sub maxp_value_Change()
    dirty = True
End Sub

Private Sub mediaOption_Click()
    mediaOption.value = True
    cartridgeOption.value = False
    current_unit% = 1
    cartridge_tester_side = 0
    'load_user_unit_stuff 1
    'setup_testvals 1
End Sub

Private Sub mediaOption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mediaOption.value = True
    cartridgeOption.value = False
    current_unit% = 1
    setup_testvals ""
End Sub

Private Sub minp_value_Change()
    dirty = True
End Sub

Private Sub testcap_Click(Index As Integer)

If Not show_advanced Then
    testcap(Index).value = 1
Else
    dirty = True
    If Index <= 15 Then '6.71.123.14
        qcshow(Index) = testcap(Index).value
    Else
        qcshow(Index + 5) = testcap(Index).value
    End If
    
End If

End Sub

Private Sub testval_Click(Index As Integer)

    Dim temp$, temp2$, POS%, fn As Integer
   
    Select Case Index
        Case 0  'outfile
            temp$ = OutFilename$(current_unit%)
           ' Me (current_unit%)
            If UCase$(OutFilename$(current_unit%)) = "NON-SELECTED" Then temp$ = ""
            fsel_name$ = temp$
            fsel_title$ = ts$(36)                   ' "Capillary Flow Output Data File"
            fsel_path$ = path(0)
            fsel_io = False
            fsel Me.hwnd
            If fsel_return$ = "" Then
                Exit Sub
            End If
            path(0) = fsel_path$
            'see if there is a period already in the string
            ' modified to allow for periods in the directory name
            POS% = RInStr(fsel_return$, "\")
            POS% = InStr(POS%, fsel_return$, ".")
            'if so, make sure it is ".cft" at end
            If POS% <> 0 Then
                'if there are not 3 characters after the period, remove everything including and
                'after the period and add ".cft"
                If (Len(fsel_return$) - POS% <> 3) Then
                    If POS% = 1 Then
                        fsel_return$ = ""
                    Else
                        fsel_return$ = Left$(fsel_return$, POS% - 1) & ".cft"
                    End If
                'if the last 3 chars are not 'cft', then "make it so"-Picard
                ElseIf (Right$(fsel_return$, 4) <> ".cft") Then
                    fsel_return$ = Left$(fsel_return$, Len(fsel_return$) - 3) & "cft"
                End If
            End If
            If cartridge_tester Then
                If mediaOption.value Then
                    current_unit% = 1
                Else
                    current_unit% = 2
                End If
                OutFilename$(current_unit%) = UCase$(fsel_return$)
            Else
                OutFilename$(current_unit%) = UCase$(fsel_return$)
            End If
            testval(0).Caption = make_sure(shorten(OutFilename$(current_unit%)))
            testval(0).SetFocus
            dirty = True
        Case 1  'line1
            GetValue.Label1.Caption = ts$(37) + ":"         ' "Enter End User"
            GetValue.Text1.Text = Line1$(current_unit%)
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(Line1$(current_unit%))
            GetValue.Label1.Tag = "text"
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                Line1$(current_unit%) = Got_Text
                dirty = True
            End If
            testval(1).Caption = make_sure(Line1$(current_unit%))
            testval(1).SetFocus
        Case 2  'line2
            GetValue.Label1.Caption = ts$(38) + ":"         ' "Enter Test Reference"
            GetValue.Text1.Text = Line2$(current_unit%)
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(Line2$(current_unit%))
            GetValue.Label1.Tag = "text"
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                Line2$(current_unit%) = Got_Text
                dirty = True
            End If
            testval(2).Caption = make_sure(Line2$(current_unit%))
            testval(2).SetFocus
        Case 3  'sid
            GetValue.Label1.Caption = ts$(39) + ":"         ' "Enter Sample ID"
            GetValue.Text1.Text = sid$(current_unit%)
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(sid$(current_unit%))
            GetValue.Label1.Tag = "text"
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                sid$(current_unit%) = Got_Text
                dirty = True
            End If
            testval(3).Caption = make_sure(sid$(current_unit%))
            testval(3).SetFocus
        Case 4 ' Lot Number
            GetValue.Label1.Caption = ts$(40) + ":"         ' "Enter Lot_Number"
            GetValue.Text1.Text = lot_number$(current_unit%)
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(lot_number$(current_unit%))
            GetValue.Label1.Tag = "text"
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                lot_number$(current_unit%) = Got_Text
                dirty = True
            End If
            testval(4).Caption = make_sure(lot_number$(current_unit%))
            testval(4).SetFocus
        Case 5 ' Operator
            GetValue.Label1.Caption = ts$(41) + ":"         ' "Enter Operator"
            GetValue.Text1.Text = operator$(current_unit%)
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(operator$(current_unit%))
            GetValue.Label1.Tag = "text"
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                operator$(current_unit%) = Got_Text
                dirty = True
            End If
            testval(5).Caption = make_sure(operator$(current_unit%))
            testval(5).SetFocus
        Case 6  'fluid, gas
            If TType%(current_unit%) = 3 Or TType%(current_unit%) = 5 Or (TType%(current_unit%) = 6 And Not lvperm_enable) Then
                NeedLiquid = False
                NeedGas = True
                NeedFluid = False
                selected_gas_or_liquid = construct_pgl(Gas$)
                GLSel.Show vbModal, Me
                Gas$ = deconstruct_pgl(selected_gas_or_liquid)
                If Len(Gas$) < 21 Then
                    testval(6).Caption = make_sure(Gas$)
                Else
                    testval(6).Caption = make_sure(Trim(Mid$(Gas$, 21)) + "   mu=" + str$(val(Left$(Gas$, 20))))
                End If
                dirty = True
                testval(6).SetFocus
            End If
            If TType%(current_unit%) = 4 Then
                NeedLiquid = True
                NeedGas = False
                NeedFluid = False
                selected_gas_or_liquid = construct_pgl(Liquid$(current_unit%))
                GLSel.Show vbModal, Me
                Liquid$(current_unit%) = deconstruct_pgl(selected_gas_or_liquid)
                If Len(Liquid$(current_unit%)) < 21 Then
                    testval(6).Caption = make_sure(Liquid$(current_unit%))
                Else
                    testval(6).Caption = make_sure(Trim(Mid$(Liquid$(current_unit%), 21)) + " mu=" + str$(val(Left$(Liquid$(current_unit%), 20))))
                End If
                dirty = True
                testval(7).Visible = False
                testval(6).SetFocus
            End If
            If TType%(current_unit%) = 1 Or TType%(current_unit%) = 2 Or (TType%(current_unit%) = 6 And lvperm_enable) Then
                NeedLiquid = False
                NeedGas = False
                NeedFluid = True
                selected_gas_or_liquid.type = epgl_other
                selected_gas_or_liquid.name = fluid$(current_unit%)
                selected_gas_or_liquid.surfTen = surfTen(current_unit%)
                GLSel.Show vbModal, Me
                fluid$(current_unit%) = selected_gas_or_liquid.name
                surfTen(current_unit%) = selected_gas_or_liquid.surfTen
                testval(6).Caption = make_sure(fluid$(current_unit%))
                testval(7).Caption = str$(surfTen(current_unit%)) + " " + ts$(44)   ' "dynes/cm"
                testval(6).SetFocus
            End If
        Case 7  'SURFTEN(current_unit%)
            NeedLiquid = False
            NeedGas = False
            NeedFluid = True
            selected_gas_or_liquid.type = epgl_other
            selected_gas_or_liquid.name = fluid$(current_unit%)
            selected_gas_or_liquid.surfTen = surfTen(current_unit%)
            GLSel.Show vbModal, Me
            fluid$(current_unit%) = selected_gas_or_liquid.name
            surfTen(current_unit%) = selected_gas_or_liquid.surfTen
            testval(6).Caption = make_sure(fluid$(current_unit%))
            testval(7).Caption = str$(surfTen(current_unit%)) + " " + ts$(44)   ' "dynes/cm"
            testval(6).SetFocus
        Case 8  'dry or testing parameters
            If TType%(current_unit%) = 2 Then Exit Sub
            If TType%(current_unit%) = 1 And TMode%(current_unit%) = 4 Then Exit Sub
            If TType%(current_unit%) = 5 Then
                Do
                    GetValue.Label1.Caption = ts$(45) + " - " + PU$             ' "Starting Pressure (differential)"
                    GetValue.Label1.Tag = ""
                    GetValue.Text1.Text = Format$(Hold_Press(current_unit%) * PCNV, "####0.#####")
                    GetValue.Text1.SelStart = 0
                    GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
                    GetValue.Continue.default = True
                    GetValue.Show 1
                    If Got_Value <> -9 Then
                        Hold_Press(current_unit%) = Got_Value / PCNV
                        dirty = True
                    End If
                    DoEvents
                Loop While Hold_Press(current_unit%) <= 0
                testval(8).Caption = ts$(15) + "=" + Format$(Hold_Press(current_unit%) * PCNV, "####0.##### ") + PU$ + ", " + ts$(16) + "=" + str$(Hold_Time(current_unit%)) + " " + ts$(17) + ", " + ts$(18) + "=" + str$(Hold_Delay(current_unit%)) + " " + ts$(17)    ' "Pressure"/"Time"/"sec""Delay"/"sec"
                If Hold_Rate(current_unit%) > 0 Then
                    testval(8).Caption = testval(8).Caption + ", " + ts$(19) + "=" + Format$(Hold_Rate(current_unit%) * PCNV, "####0.##### ") + PU$ + IIf(PH_fail_method = "dpdt", "/" + IIf(pressHoldUnit = "S", "sec", "min"), "") ' "Fail"/"sec"
                End If
            ElseIf TType%(current_unit%) = 6 Then
                If MsgBox(ts$(50), 4, ts$(51)) = 6 Then               ' "Do you want to load an old test?"/"Diff. Perm."
                    fsel_name$ = ""
                    fsel_title$ = ts$(51) + " " + ts$(52)               ' "Diff. Perm."/"Test Method File"
                    fsel_path$ = Left$(path(1), Len(path(1)) - 3) + "DTM"
                    fsel_io = True
                    fsel Me.hwnd
                    If fsel_return$ = "" Then
                        Exit Sub
                    End If
                    fn = FreeFile
                    Open fsel_return$ For Input As #fn
                    Input #fn, mf_press(current_unit%)
                    Input #fn, mf_time(current_unit%)
                    Input #fn, Step_Time(current_unit%)
                    ' Now get the stuff that may or may not be there for the linear seal
                    On Error Resume Next
                    Input #fn, temp$
                    Input #fn, temp2$
                    If temp$ = "" Then
                        MF_linearSeal = False
                        seal_state = 0
                    ElseIf temp2$ = "" Then
                        MF_linearSeal = True
                        seal_state = 1
                        MF_sealDiam = val(temp$)
                    Else
                        MF_linearSeal = True
                        seal_state = 2
                        MF_innerDiam = val(temp$)
                        MF_outerDiam = val(temp2$)
                    End If
                    On Error GoTo 0
                    Close #fn
                    dirty = True
                Else
                    If lvperm_enable Then
                        ' no holding pressure for this
                    Else
                      Do
                        GetValue.Label1.Caption = ts$(53) + " - " + PU$     ' "Internal Pressure (differential)"
                        GetValue.Label1.Tag = ""
                        GetValue.Text1.Text = Format$(mf_press(current_unit%) * PCNV, "####0.#####")
                        GetValue.Text1.SelStart = 0
                        GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
                        GetValue.Continue.default = True
                        GetValue.Show 1
                        If Got_Value <> -9 Then
                            mf_press(current_unit%) = Got_Value / PCNV
                            dirty = True
                        End If
                        DoEvents
                      Loop While mf_press(current_unit%) <= 0
                    End If
                    Do
                        GetValue.Label1.Caption = ts$(46) + " (" + ts$(21) + ")" '   "Maximum Test Duration"  - min."
                        GetValue.Label1.Tag = ""
                        GetValue.Text1.Text = str$(mf_time(current_unit%))
                        GetValue.Text1.SelStart = 0
                        GetValue.Text1.SelLength = 10
                        GetValue.Continue.default = True
                        GetValue.Show 1
                        If Got_Value <> -9 Then
                            mf_time(current_unit%) = Got_Value
                            dirty = True
                        End If
                        DoEvents
                    Loop While mf_time(current_unit%) <= 0
                    Do
                        GetValue.Label1.Caption = ts$(54) + " (" + ts$(21) + ")" ' "Time between data points - min."
                        GetValue.Label1.Tag = ""
                        GetValue.Text1.Text = str$(Step_Time(current_unit%))
                        GetValue.Text1.SelStart = 0
                        GetValue.Text1.SelLength = 10
                        GetValue.Continue.default = True
                        GetValue.Show 1
                        If Got_Value <> -9 Then
                            Step_Time(current_unit%) = Got_Value
                            dirty = True
                        End If
                        DoEvents
                    Loop While Step_Time(current_unit%) <= 0
                    If MF_linearSeal Then
                        If seal_state = 1 Then
                            Do
                                With GetValue
                                    .Label1.Caption = ts$(70) + " - cm."        ' "Seal diameter"
                                    .Label1.Tag = ""
                                    .Text1.Text = str$(MF_sealDiam)
                                    .Text1.SelStart = 0
                                    .Text1.SelLength = 10
                                    .Continue.default = True
                                    .Show 1
                                End With
                                If Got_Value <> -9 Then
                                    MF_sealDiam = Got_Value
                                    dirty = True
                                End If
                                DoEvents
                            Loop While MF_sealDiam < 0
                        ElseIf seal_state = 2 Then
                            Do
                                With GetValue
                                    .Label1.Caption = ts$(71) + " - cm."        ' "Inner seal diameter"
                                    .Label1.Tag = ""
                                    .Text1.Text = str$(MF_innerDiam)
                                    .Text1.SelStart = 0
                                    .Text1.SelLength = 10
                                    .Continue.default = True
                                    .Show 1
                                End With
                                If Got_Value <> -9 Then
                                    MF_innerDiam = Got_Value
                                    dirty = True
                                End If
                                DoEvents
                            Loop While MF_innerDiam < 0
                            Do
                                With GetValue
                                    .Label1.Caption = ts$(72) + " - cm."        ' "Outer seal diameter"
                                    .Label1.Tag = ""
                                    .Text1.Text = str$(MF_outerDiam)
                                    .Text1.SelStart = 0
                                    .Text1.SelLength = 10
                                    .Continue.default = True
                                    .Show 1
                                End With
                                If Got_Value <> -9 Then
                                    MF_outerDiam = Got_Value
                                    dirty = True
                                End If
                                DoEvents
                            Loop While MF_outerDiam < 0
                        End If
                    End If
                    If MsgBox(ts$(55), 4, ts$(51)) = 6 Then        ' "Do you want to save this test?"/"Diff. Perm."/
                        fsel_name$ = ""
                        fsel_title$ = ts$(51) + " " + ts$(52)       ' "Diff. Perm. Test Method File"
                        fsel_path$ = Left$(path(1), Len(path(1)) - 3) + "DTM"
                        fsel_io = False
                        fsel Me.hwnd
                        If fsel_return$ <> "" Then
                            fn = FreeFile
                            Open fsel_return$ For Output As #fn
                            Print #fn, mf_press(current_unit%)
                            Print #fn, mf_time(current_unit%)
                            Print #fn, Step_Time(current_unit%)
                            If seal_state = 1 Then
                                Print #fn, MF_sealDiam
                            ElseIf seal_state = 2 Then
                                Print #fn, MF_innerDiam
                                Print #fn, MF_outerDiam
                            End If
                            Close #fn
                        End If
                    End If
                End If
                If lvperm_enable Then
                    testval(8).Caption = ts$(16) + "=" + str$(mf_time(current_unit%)) + " " + ts$(21) + ", " + ts$(20) + " " + ts$(16) + "=" + str$(Step_Time(current_unit%)) + " " + ts$(21) ' "Time"/"min"/"Data"/"Time"/"min"
                Else
                    set_micro_text
                End If
            Else
              fsel_name$ = ""
              fsel_title$ = ts$(56) + " " + ts$(3) + " " + ts$(57) ' "Capillary Flow DRY Parameter File."
              fsel_path$ = path(1)
              fsel_io = True
              fsel Me.hwnd
              If fsel_return$ = "" Then
                Exit Sub
              End If
              path(1) = fsel_path$
              TPFDRY$(current_unit%) = UCase$(fsel_return$)
              If testval(8).Caption = testval(9).Caption Then
                TPFWET$(current_unit%) = TPFDRY$(current_unit%)
                testval(9).Caption = make_sure(shorten(TPFWET$(current_unit%)))
              End If
              testval(8).Caption = make_sure(shorten(TPFDRY$(current_unit%)))
              dirty = True
            End If
            testval(8).SetFocus
        Case 9  'wet
            reply% = 0
            If TType%(current_unit%) = 1 And TMode%(current_unit%) < 4 Then
                reply% = MsgBox(ts$(58), 4, ts$(29))       ' "Would you like to select a WET Parameter File that is different than the DRY?"/"Test Setup"
            End If
            If reply% = 7 Then
                TPFWET$(current_unit%) = UCase$(TPFDRY$(current_unit%))
            Else
                fsel_name$ = ""
                fsel_title$ = ts$(56) + " " + ts$(6) + " " + ts$(57) ' "Capillary Flow WET Parameter File."
                fsel_path$ = path(1)
                fsel_io = True
                fsel Me.hwnd
                If fsel_return$ = "" Then
                    Exit Sub
                End If
                path(1) = fsel_path$
                TPFWET$(current_unit%) = UCase$(fsel_return$)
                dirty = True
            End If
            testval(9).Caption = make_sure(shorten(TPFWET$(current_unit%)))
            testval(9).SetFocus
        Case 10  'ttype

            If Not WESA_exclusive Then
                'T_Select$ = ""
                'Selection.Show 1
                'Refresh
                'If T_Select$ <> "Cancel" Then
                
                Select_Test.Show 1
                testval(10).Caption = TestName$()
                testval(11).Caption = Format$(Diam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
                setup_testname
                dirty = True
            End If
            
            ' Check to see if pressure/pore diameter units have to be switched
            If minmaxunits = "d" Then           ' no problem for pressure
                If TType%(current_unit%) >= 3 And p_units_label.Caption = "microns" Then
                    ' Don't have surface tension, so switch to pressure
                    Label4.Caption = get_thing("test", "Label4", language$, Label4.Caption, Label4, default_font)
                    Label5.Caption = get_thing("test", "Label5", language$, Label5.Caption, Label5, default_font)
                    p_units_label = PU$
                ElseIf TType%(current_unit%) < 3 Then
                    ' Not in diameter, but should be
                    Label4.Caption = ts$(68)
                    Label5.Caption = ts$(69)
                    p_units_label = "microns"
               End If
            End If
                    
           ' testval(10).SetFocus
        Case 11 'Diam
            If use_fixed_sample_diameter_cm Then '6.71.123.14
                MsgBox ts$(74)                                  ' "This unit has a fixed outlet diameter."
            Else
                'If numberOfMicrometers > 0 Then
                '    readMicrometer currentMicrometer
                '    Diam(current_unit%) = micrometers(currentMicrometer).dblValue
                '    testval(11).Caption = Format$(Diam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
                'Else
                If TType%(current_unit%) = 1 And CF_SampleType = 1 Then
                    GetValue.Label1.Caption = "Enter Inner Diameter " + " (" + linear_unit_name$ + "):"     ' "Enter Diameter"
                    GetValue.Text1.Text = Format$(innerDiam(current_unit%) / linear_unit_conversion#, "####0.#####")
                    GetValue.Text1.SelStart = 0
                    GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
                    GetValue.Continue.default = True
                    GetValue.Show 1
                    If Got_Value > 0 Then
                        innerDiam(current_unit%) = Got_Value * linear_unit_conversion#
                        dirty = True
                    End If
                    testval(11).Caption = Format$(innerDiam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
                Else
                    GetValue.Label1.Caption = ts$(59) + " (" + linear_unit_name$ + "):"     ' "Enter Diameter"
                    GetValue.Text1.Text = Format$(Diam(current_unit%) / linear_unit_conversion#, "####0.#####")
                    GetValue.Text1.SelStart = 0
                    GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
                    GetValue.Continue.default = True
                    GetValue.Show 1
                    If Got_Value > 0 Then
                        Diam(current_unit%) = Got_Value * linear_unit_conversion#
                        dirty = True
                    End If
                    testval(11).Caption = Format$(Diam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
                End If
                'End If
                testval(11).SetFocus
            End If
        Case 12 'Cyl_Len(current_unit%)
            If TType%(current_unit%) = 1 And CF_SampleType = 1 Then
                GetValue.Label1.Caption = "Enter Outer Diameter " + " (" + linear_unit_name$ + "):"     ' "Enter Cyl. Length"
                GetValue.Text1.Text = Format$(outerDiam(current_unit%) / linear_unit_conversion#, "####0.#####")
                GetValue.Text1.SelStart = 0
                GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
                GetValue.Continue.default = True
                GetValue.Show 1
                If Got_Value >= 0 Then
                    outerDiam(current_unit%) = Got_Value * linear_unit_conversion#
                    dirty = True
                End If
                testval(12).Caption = Format$(outerDiam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
            Else
                GetValue.Label1.Caption = ts$(60) + " (" + linear_unit_name$ + "):"     ' "Enter Cyl. Length"
                GetValue.Text1.Text = Format$(cyl_len(current_unit%) / linear_unit_conversion#, "####0.#####")
                GetValue.Text1.SelStart = 0
                GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
                GetValue.Continue.default = True
                GetValue.Show 1
                If Got_Value >= 0 Then
                    cyl_len(current_unit%) = Got_Value * linear_unit_conversion#
                    dirty = True
                End If
                testval(12).Caption = Format$(cyl_len(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
            End If
            testval(12).SetFocus
            If TType%(current_unit%) = 6 And lvperm_enable = False Then
                check_micro_seal_type
                set_micro_text
            End If
        Case 13 'thick(current_unit%)
            If numberOfMicrometers > 0 Then
                readMicrometer currentMicrometer
                thick(current_unit%) = micrometers(currentMicrometer).dblValue
                testval(13).Caption = Format$(Diam(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
            End If
            
            If TType%(current_unit%) = 1 And CF_SampleType = 1 Then
                GetValue.Label1.Caption = "Enter Sample Height " + " (" + linear_unit_name$ + "):"      ' "Enter Thickness"
                GetValue.Text1.Text = Format$(cyl_len(current_unit%) / linear_unit_conversion#, "####0.#####")
                GetValue.Text1.SelStart = 0
                GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
                GetValue.Continue.default = True
                GetValue.Show 1
                If Got_Value > 0 Then
                    cyl_len(current_unit%) = Got_Value * linear_unit_conversion#
                    dirty = True
                End If
                testval(13).Caption = Format$(cyl_len(current_unit%) / linear_unit_conversion#, "####0.##### ") + linear_unit_name$
            Else
                GetValue.Label1.Caption = ts$(61) + " (" + thick_unit_name$ + "):"      ' "Enter Thickness"
                GetValue.Text1.Text = Format$(thick(current_unit%) / thick_unit_conversion#, "####0.#####")
                GetValue.Text1.SelStart = 0
                GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
                GetValue.Continue.default = True
                GetValue.Show 1
                If Got_Value > 0 Then
                    thick(current_unit%) = Got_Value * thick_unit_conversion#
                    dirty = True
                End If
                testval(13).Caption = Format$(thick(current_unit%) / thick_unit_conversion#, "####0.##### ") + thick_unit_name$
            End If
            testval(13).SetFocus
        Case 14
            GetValue.Label1.Caption = ts$(62) + " " + dens_unit + ": "      ' "Enter absolute density in"
            GetValue.Text1.Text = Format$(density(current_unit%) / dens_unit_c, "####0.#####")
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value > 0 Then
                density(current_unit%) = Got_Value * dens_unit_c
                dirty = True
            End If
            testval(14).Caption = Format$(density(current_unit%) / dens_unit_c, "####0.#####") + dens_unit
            testval(14).SetFocus
        Case 15
            GetValue.Label1.Caption = ts$(63) + " " + mass_unit + ": "      ' "Enter mass in"
            GetValue.Text1.Text = Format$(mass(current_unit%) / mass_unit_c, "####0.#####")
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value > 0 Then
                mass(current_unit%) = Got_Value * mass_unit_c
                dirty = True
            End If
            If Not use_fixed_sample_diameter_cm Then '6.71.123.14
                GetValue.Label1.Caption = ts$(64) + " " + linear_unit_name + ":"
                GetValue.Text1.Text = Format$(SCDiam / linear_unit_conversion, "####0.#####")
                GetValue.Text1.SelStart = 0
                GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
                GetValue.Continue.default = True
                GetValue.Show 1
                GetValue.Label1.Tag = ""
                If Got_Value <> -9 Then
                    SCDiam = Got_Value * linear_unit_conversion
                    dirty = True
                End If
            End If
            If SCDiam > Diam(current_unit%) Then
                testval(15).Caption = Format$(mass(current_unit%) / mass_unit_c, "####0.#####") + mass_unit + ";  " + ts$(65) + ": " + Format$(SCDiam / linear_unit_conversion#, "####0.##### ") + linear_unit_name$    ' "Sample Cut Diameter for Measured Mass"
            Else
                testval(15).Caption = Format$(mass(current_unit%) / mass_unit_c, "####0.#####") + mass_unit
            End If
            testval(15).SetFocus
        Case 16 '6.71.123.14
            '% Porosity
            GetValue.Label1.Caption = ts$(75) + ": "      ' "Enter % Porosity"
            GetValue.Text1.Text = Format$(TargetPercPorosity, "####0.#####")
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                'no units to convert
                TargetPercPorosity = Got_Value
                dirty = True
            End If
            testval(16).Caption = Format$(TargetPercPorosity, "####0.#####")
            testval(16).SetFocus
        Case 17 '6.71.123.14
            GetValue.Label1.Caption = ts$(76) + ": "      ' "Enter Build Cake By Pressure Target Pressure"
            GetValue.Text1.Text = Format$(BuildCakeByPressureTargetPressure * PCNV, "####0.#####") + " " + PU$
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                BuildCakeByPressureTargetPressure = Got_Value / PCNV
                dirty = True
            End If
            testval(17).Caption = Format$(BuildCakeByPressureTargetPressure, "####0.#####")
            testval(17).SetFocus
        Case 18 '6.71.123.14
            GetValue.Label1.Caption = ts$(77) + ": "      ' "Enter Build Cake By Flow Target Flow"
            GetValue.Text1.Text = Format$(BuildCakeByFlowTargetFlow, "####0.#")
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                BuildCakeByFlowTargetFlow = Got_Value
                dirty = True
            End If
            testval(18).Caption = Format$(BuildCakeByFlowTargetFlow, "####0.#")
            testval(18).SetFocus
        Case 19 '6.71.123.14
            GetValue.Label1.Caption = ts$(78) + ": "      ' "Enter Build Cake By Flow End Pressure"
            GetValue.Text1.Text = Format$(BuildCakeByFlowEndPressure * PCNV, "####0.#") + " " + PU$
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                BuildCakeByFlowEndPressure = Got_Value / PCNV
                dirty = True
            End If
            testval(19).Caption = Format$(BuildCakeByFlowEndPressure, "####0.#####")
            testval(19).SetFocus
        Case 20 '6.71.123.14
            GetValue.Label1.Caption = ts$(79) + ": "      ' "Enter Slurry Tube Wash Cycle Target Flow"
            GetValue.Text1.Text = Format$(SlurryTubeWashCycleTargetFlow, "####0.#")
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                SlurryTubeWashCycleTargetFlow = Got_Value
                dirty = True
            End If
            testval(20).Caption = Format$(SlurryTubeWashCycleTargetFlow, "####0.#")
            testval(20).SetFocus
        Case 21
            GetValue.Label1.Caption = ts$(83) + ": "      ' "Enter Slurry Tube Wash Cycle Target Flow"
            GetValue.Text1.Text = Format$(BuildCakeByPressureEndFlow, "####0.#")
            GetValue.Text1.SelStart = 0
            GetValue.Text1.SelLength = Len(GetValue.Text1.Text)
            GetValue.Continue.default = True
            GetValue.Show 1
            GetValue.Label1.Tag = ""
            If Got_Value <> -9 Then
                BuildCakeByPressureEndFlow = Got_Value
                dirty = True
            End If
            testval(21).Caption = Format$(BuildCakeByPressureEndFlow, "####0.#")
            testval(21).SetFocus
        End Select
    Refresh
    
End Sub

Function make_sure(ByVal a$) As String

Dim i As Integer, j As Integer

If Len(a$) = 0 Then
    make_sure = " "
Else
    ' search out all ampersands in the string and double them
    i = 1
    Do
        j = InStr(i, a$, "&")
        If j = 0 Then Exit Do
        ' there is an ampersand at position j in a$
        a$ = Left$(a$, j) & Mid$(a$, j)
        i = j + 2
    Loop
    make_sure = a$
End If

End Function
Private Sub testval_GotFocus(Index As Integer)
    Me.Tag = str$(Index)
End Sub

Private Sub tf_check_Click()

If tf_check.value = 1 Then
  tf_check.Caption = ts$(66) + ":"              ' "User Defined"
  tf_value.Visible = True
Else
  tf_check.Caption = ts$(67) + " (0.715)"       ' "Default"
  tf_value.Visible = False
End If

End Sub

Private Sub tf_value_LostFocus()

If myVal(tf_value.Text) <= 0 Then
    tf_check.value = 0
    tf_value.Text = "0.715"
End If

End Sub
Private Sub update_sample_ID()
    '============= New Code edc 05-04-07============================================
   ' Dim lastSlash, next_Slash As Long
   ' Dim srtSampID As String
    
   ' If AutoSamplID Then
   ' lastSlash = InStrRev(OutFilename(current_unit), "\")
   ' next_Slash = InStrRev(OutFilename(current_unit), "\", lastSlash - 1)
   ' srtSampID = Mid(OutFilename(current_unit), next_Slash, (lastSlash - next_Slash))
   '
   ' End If
    '============= End New Code=====================================================
End Sub
Private Sub update_filename()

    Dim alldigits As Boolean
    Dim a$, i As Integer

    If Not auto_increment Then Exit Sub
    Rem Check last three characters of file name (ignore .cft ending)
    ' if they are all from 0 to 9, increment this number by 1
    If Len(OutFilename$(current_unit%)) >= 7 Then
        a$ = Left$(Right$(OutFilename$(current_unit%), 7), 3)
        alldigits = True
        For i = 1 To 3
            If Mid$(a$, i, 1) < "0" Or Mid$(a$, i, 1) > "9" Then
                alldigits = False
                Exit For
            End If
        Next i
        If alldigits Then
            i = Len(OutFilename$(current_unit%)) - 4
            Do
                a$ = Mid$(OutFilename$(current_unit%), i, 1)
                If a$ >= "0" And a$ <= "8" Then
                    a$ = Chr$(Asc(a$) + 1)
                    Mid$(OutFilename$(current_unit%), i, 1) = a$
                    Exit Do
                ElseIf a$ <> "9" Or i = 1 Then
                    a$ = "1" + Mid$(OutFilename$(current_unit%), i + 1)
                    OutFilename$(current_unit%) = Left$(OutFilename$(current_unit%), i) + a$
                    Exit Do
                End If
                ' digit is a 9 - need to do next one to left
                Mid$(OutFilename$(current_unit%), i, 1) = "0"
                i = i - 1
            Loop
        End If
    End If
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Continue_Click
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    
    Testscrn.Caption = get_thing("test", "window title", language$, Testscrn.Caption, Testscrn, default_font)
    Cancelit.Caption = gpps2("test", "cancel", language$, Cancelit.Caption)
    set_fontname Cancelit, default_font
    cmdChgUser.Caption = gpps2("test", "chguser", language$, cmdChgUser.Caption)
    set_fontname cmdChgUser, default_font
    cmdEditParamFile.Caption = gpps2("test", "editparamfile", language$, cmdEditParamFile.Caption)
    set_fontname cmdEditParamFile, default_font
    cmdEditPreferences.Caption = gpps2("test", "editprefs", language$, cmdEditPreferences.Caption)
    set_fontname cmdEditPreferences, default_font
    cmdPoreShape.Caption = gpps2("test", "poreshape", language$, cmdPoreShape.Caption)
    set_fontname cmdPoreShape, default_font
    Command1.Caption = gpps2("test", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    Continue.Caption = gpps2("test", "continue", language$, Continue.Caption)
    set_fontname Continue, default_font
    Help.Caption = gpps2("test", "help", language$, Help.Caption)
    set_fontname Help, default_font
    For i = 0 To 3
        Label1(i).Caption = get_thing("test", "Label1" + str$(i), language$, Label1(i).Caption, Label1(i), default_font)
    Next i
    Label4.Caption = get_thing("test", "Label4", language$, Label4.Caption, Label4, default_font)
    Label5.Caption = get_thing("test", "Label5", language$, Label5.Caption, Label5, default_font)
    set_fontstuff minp_value, default_font
    set_fontstuff maxp_value, default_font
    set_fontstuff p_units_label, default_font
    '6.71.123.14 For i = 0 To 15
    
    For i = 0 To 15
        testcap(i).Caption = get_thing("test", "testcap" + str$(i), language$, testcap(i).Caption, testcap(i), default_font)
        Dim s As String
        s = testcap(i).Caption
        testval(i).Caption = get_thing("test", "testval" + str$(i), language$, testval(i).Caption, testval(i), default_font)
    Next i
    tf_check.Caption = get_thing("test", "tfcheck", language$, tf_check.Caption, tf_check, default_font)
    tf_label.Caption = get_thing("test", "tflabel", language$, tf_label.Caption, tf_label, default_font)
    set_fontstuff tf_value, default_font
    
    set_fontstuff Label2, default_font
 
    ' Other text
    ts$(1) = gpps2("test", "ts1", language$, "Do you want to save any changes you made?")
    ts$(2) = gpps2("test", "ts2", language$, "Pressure in")
    ts$(3) = gpps2("test", "ts3", language$, "Dry Parameter")
    ts$(4) = gpps2("test", "ts4", language$, "Sample cut diameter for measured mass")
    ts$(5) = gpps2("test", "ts5", language$, "Fluid")
    ts$(6) = gpps2("test", "ts6", language$, "Wet Parameter")
    ts$(7) = gpps2("test", "ts7", language$, "Warning: the")
    ts$(8) = gpps2("test", "ts8", language$, "file is invalid.")
    ts$(9) = gpps2("test", "ts9", language$, "Gas")
    ts$(10) = gpps2("test", "ts10", language$, "Gas Permeability")
    ts$(11) = gpps2("test", "ts11", language$, "Envelope Surface Area Analysis")
    ts$(12) = gpps2("test", "ts12", language$, "Pressure Drop Test")
    ts$(13) = gpps2("test", "ts13", language$, "Liquid")
    ts$(14) = gpps2("test", "ts14", language$, "Hold Parameters")
    ts$(15) = gpps2("test", "ts15", language$, "Pressure")
    ts$(16) = gpps2("test", "ts16", language$, "Time")
    ts$(17) = gpps2("test", "ts17", language$, "sec")
    ts$(18) = gpps2("test", "ts18", language$, "Delay")
    ts$(19) = gpps2("test", "ts19", language$, "Fail")
    ts$(20) = gpps2("test", "ts20", language$, "Data")
    ts$(21) = gpps2("test", "ts21", language$, "min")
    ts$(22) = gpps2("test", "ts22", language$, "&Pore Shape: Circular")
    ts$(23) = gpps2("test", "ts23", language$, "&Pore Shape: Square")
    ts$(24) = gpps2("test", "ts24", language$, "Your instrument does not have hydrohead test capabilities.")
    ts$(25) = gpps2("test", "ts25", language$, "Please check.")
    ts$(26) = gpps2("test", "ts26", language$, "Your instrument does not have envelope surface area test capabilities.")
    ts$(27) = gpps2("test", "ts27", language$, "Your instrument does not have pressure drop test capabilities.")
    ts$(28) = gpps2("test", "ts28", language$, "The selected file already exists. Overwrite?")
    ts$(29) = gpps2("test", "ts29", language$, "Test Setup")
    ts$(30) = gpps2("test", "ts30", language$, "Maximum pressure must be greater than minimum pressure")
    ts$(31) = gpps2("test", "ts31", language$, "Warning: I thought there were more chambers to setup, but I can't find any.  Please let PMI know that you have seen this message.")
    ts$(32) = gpps2("test", "ts32", language$, "Auto Test Setup: Chamber")
    ts$(33) = gpps2("test", "ts33", language$, "&Start Test")
    ts$(34) = gpps2("test", "ts34", language$, "&Next Unit")
    ts$(35) = gpps2("test", "ts35", language$, "Unit")
    ts$(36) = gpps2("test", "ts36", language$, "Capillary Flow Output Data File")
    ts$(37) = gpps2("test", "ts37", language$, "Enter End User")
    ts$(38) = gpps2("test", "ts38", language$, "Enter Test Reference")
    ts$(39) = gpps2("test", "ts39", language$, "Enter Sample ID")
    ts$(40) = gpps2("test", "ts40", language$, "Enter Lot_Number")
    ts$(41) = gpps2("test", "ts41", language$, "Enter Operator")
    ts$(42) = gpps2("test", "ts42", language$, "Enter Fluid name")
    ts$(43) = gpps2("test", "ts43", language$, "Enter surface tension")
    ts$(44) = gpps2("test", "ts44", language$, "dynes/cm")
    ts$(45) = gpps2("test", "ts45", language$, "Starting pressure (differential)")
    ts$(46) = gpps2("test", "ts46", language$, "Maximum test duration")
    ts$(47) = gpps2("test", "ts47", language$, "Initial delay")
    ts$(48) = gpps2("test", "ts48", language$, "Fail pressure rate")
    ts$(49) = gpps2("test", "ts49", language$, "or 0 for none")
    ts$(50) = gpps2("test", "ts50", language$, "Do you want to load an old test?")
    ts$(51) = gpps2("test", "ts51", language$, "Diff. Perm.")
    ts$(52) = gpps2("test", "ts52", language$, "Test Method File")
    ts$(53) = gpps2("test", "ts53", language$, "Internal pressure (differential)")
    ts$(54) = gpps2("test", "ts54", language$, "Time between data points")
    ts$(55) = gpps2("test", "ts55", language$, "Do you want to save this test?")
    ts$(56) = gpps2("test", "ts56", language$, "Capillary Flow")
    ts$(57) = gpps2("test", "ts57", language$, "File")
    ts$(58) = gpps2("test", "ts58", language$, "Would you like to select a Wet parameter file that is different than the Dry?")
    ts$(59) = gpps2("test", "ts59", language$, "Enter Diameter")
    ts$(60) = gpps2("test", "ts60", language$, "Enter Cyl. Length")
    ts$(61) = gpps2("test", "ts61", language$, "Enter Thickness")
    ts$(62) = gpps2("test", "ts62", language$, "Enter absolute density in")
    ts$(63) = gpps2("test", "ts63", language$, "Enter mass in")
    ts$(64) = gpps2("test", "ts64", language$, "Enter sample diameter for measured mass in")
    ts$(65) = gpps2("test", "ts65", language$, "Sample cut diameter for measured mass")
    ts$(66) = gpps2("test", "ts66", language$, "User Defined")
    ts$(67) = gpps2("test", "ts67", language$, "Default")
    ts$(68) = gpps2("test", "ts68", language$, "Minimum pore diameter")
    ts$(69) = gpps2("test", "ts69", language$, "Maximum pore diameter")
    ts$(70) = gpps2("test", "ts70", language$, "Seal diameter")
    ts$(71) = gpps2("test", "ts71", language$, "Inner seal diameter")
    ts$(72) = gpps2("test", "ts72", language$, "Outer seal diameter")
    ts$(73) = gpps2("test", "ts73", language$, "Maximum pressure is too high for safe operation with the current compression pressure")
    ts$(74) = gpps2("test", "ts74", language$, "This unit has a fixed outlet diameter.") '6.71.123.14
    ts$(75) = gpps2("test", "ts75", language$, "Enter % Porosity") '6.71.123.14
    ts$(76) = gpps2("test", "ts76", language$, "Enter Build Cake By Pressure Target Pressure") '6.71.123.14
    ts$(77) = gpps2("test", "ts77", language$, "Enter Build Cake By Flow Target Flow") '6.71.123.14
    ts$(78) = gpps2("test", "ts78", language$, "Enter Build Cake By Flow End Pressure") '6.71.123.14
    ts$(79) = gpps2("test", "ts79", language$, "Enter Slurry Tube Wash Cycle Target Flow") '6.71.123.14
    ts$(80) = gpps2("test", "ts80", language$, "Gurley Method") '6.71.123.30
    ts$(81) = gpps2("test", "ts81", language$, "High Pressure Gurley") '6.71.123.30
    ts$(82) = gpps2("test", "ts82", language$, "Sheffield Method") '6.71.123.30
    ts$(83) = gpps2("test", "ts83", language$, "Enter Build Cake By Pressure End Flow")
    ts$(84) = gpps2("test", "ts84", language$, "Frazier Test")
End Sub
