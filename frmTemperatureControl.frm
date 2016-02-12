VERSION 5.00
Begin VB.Form frmTemperatureControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temperature Control"
   ClientHeight    =   5985
   ClientLeft      =   5955
   ClientTop       =   5025
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6495
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveSettings 
      Caption         =   "Save Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame frameTemperatureControl 
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6015
      Begin VB.CheckBox chkDelay 
         Height          =   255
         Index           =   7
         Left            =   5280
         TabIndex        =   37
         Top             =   3960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkSet 
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   36
         Top             =   3960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtArray 
         Enabled         =   0   'False
         Height          =   375
         Index           =   7
         Left            =   2880
         TabIndex        =   35
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkDelay 
         Height          =   255
         Index           =   6
         Left            =   5280
         TabIndex        =   33
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkSet 
         Height          =   255
         Index           =   6
         Left            =   4560
         TabIndex        =   32
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtArray 
         Enabled         =   0   'False
         Height          =   375
         Index           =   6
         Left            =   2880
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkDelay 
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   29
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkDelay 
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   28
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkDelay 
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   27
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkDelay 
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkDelay 
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkDelay 
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkSet 
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   23
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkSet 
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkSet 
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkSet 
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkSet 
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkSet 
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtArray 
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   2880
         TabIndex        =   8
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArray 
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   2880
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArray 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   2880
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArray 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArray 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArray 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblArray 
         Caption         =   "Mullen Temperature"
         Enabled         =   0   'False
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
         Index           =   7
         Left            =   240
         TabIndex        =   34
         Top             =   4020
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblArray 
         Caption         =   "Hydro Head Temperature"
         Enabled         =   0   'False
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
         Index           =   6
         Left            =   240
         TabIndex        =   30
         Top             =   3540
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblDelay 
         Caption         =   "Delay?"
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
         Left            =   5160
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblSet 
         Caption         =   "Set?"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblArray 
         Caption         =   "Cabinet Temperature"
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   3060
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblArray 
         Caption         =   "Bubbler Temperature"
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   2580
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblArray 
         Caption         =   "Air Temperature"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   2100
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblArray 
         Caption         =   "Reservoir Temperature"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1620
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblArray 
         Caption         =   "Wet Chamber Temperature"
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   1140
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblArray 
         Caption         =   "Dry Chamber Temperature"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.CheckBox chkUseTempControl 
      Caption         =   "Use Temperature Control"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmTemperatureControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUseTempControl_Click()
    Dim i As Integer
    Dim b As Boolean
    
    b = IIf(chkUseTempControl.value = 1, True, False)
    For i = 0 To 7
        lblArray(i).Enabled = b
        txtArray(i).Enabled = b
        chkSet(i).Enabled = b
        chkDelay(i).Enabled = b
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSaveSettings_Click()
    Dim u$
    Dim i As Integer
    
    If dryChamberTemperature <> 0 Then
        dryChamberTargetTemperature(current_unit%) = verifyValidValue(val(txtArray(0).Text))
    End If
    If wetChamberTemperature <> 0 Then
        wetChamberTargetTemperature(current_unit%) = verifyValidValue(val(txtArray(1).Text))
    End If
    If reservoirTemperature <> 0 Then
        reservoirTargetTemperature(current_unit%) = verifyValidValue(val(txtArray(2).Text))
    End If
    If airTemperature <> 0 Then
        airTargetTemperature(current_unit%) = verifyValidValue(val(txtArray(3).Text))
    End If
    If bubblerTemperature <> 0 Then
        bubblerTargetTemperature(current_unit%) = verifyValidValue(val(txtArray(4).Text))
    End If
    If cabinetTemperature <> 0 Then
        cabinetTargetTemperature(current_unit%) = verifyValidValue(val(txtArray(5).Text))
    End If
    If hydroHeadTemperature <> 0 Then
        hydroHeadTargetTemperature(current_unit%) = verifyValidValue(val(txtArray(6).Text))
    End If
    If mullenTemperature <> 0 Then
        mullenTargetTemperature(current_unit%) = verifyValidValue(val(txtArray(7).Text))
    End If
    
    For i = 0 To 7
        setTemperatureForAuto(current_unit%, i) = IIf(chkSet(i).value = 1, True, False)
        delayTestForTemperature(current_unit%, i) = IIf(chkDelay(i).value = 1, True, False)
    Next i
    
    useTemperatureControlForAuto(current_unit%) = IIf(chkUseTempControl.value = 0, False, True)
    If current_unit% = 1 Then u$ = "" Else u$ = Format$(current_unit%)
    save_user_stuff u$
    Unload Me
End Sub

Private Function verifyValidValue(testValue As Single)
    If testValue < minimumPossibleTemperature Then
        verifyValidValue = minimumPossibleTemperature
    ElseIf testValue > maximumPossibleTemperature Then
        verifyValidValue = maximumPossibleTemperature
    Else
        verifyValidValue = testValue
    End If
End Function


Private Sub Form_Load()
    Dim i As Integer
    
    If dryChamberTemperature <> 0 Then
        lblArray(0).Visible = True
        txtArray(0).Visible = True
        txtArray(0).Text = dryChamberTargetTemperature(current_unit%)
        chkSet(0).Visible = True
        chkDelay(0).Visible = True
    End If
    If wetChamberTemperature <> 0 Then
        lblArray(1).Visible = True
        txtArray(1).Visible = True
        txtArray(1).Text = wetChamberTargetTemperature(current_unit%)
        chkSet(1).Visible = True
        chkDelay(1).Visible = True
    End If
    If reservoirTemperature <> 0 Then
        lblArray(2).Visible = True
        txtArray(2).Visible = True
        txtArray(2).Text = reservoirTargetTemperature(current_unit%)
        chkSet(2).Visible = True
        chkDelay(2).Visible = True
    End If
    If airTemperature <> 0 Then
        lblArray(3).Visible = True
        txtArray(3).Visible = True
        txtArray(3).Text = airTargetTemperature(current_unit%)
        chkSet(3).Visible = True
        chkDelay(3).Visible = True
    End If
    If bubblerTemperature <> 0 Then
        lblArray(4).Visible = True
        txtArray(4).Visible = True
        txtArray(4).Text = bubblerTargetTemperature(current_unit%)
        chkSet(4).Visible = True
        chkDelay(4).Visible = True
    End If
    If cabinetTemperature <> 0 Then
        lblArray(5).Visible = True
        txtArray(5).Visible = True
        txtArray(5).Text = cabinetTargetTemperature(current_unit%)
        chkSet(5).Visible = True
        chkDelay(5).Visible = True
    End If
    If hydroHeadTemperature <> 0 Then
        lblArray(6).Visible = True
        txtArray(6).Visible = True
        txtArray(6).Text = hydroHeadTargetTemperature(current_unit%)
        chkSet(6).Visible = True
        chkDelay(6).Visible = True
    End If
    If mullenTemperature <> 0 Then
        lblArray(7).Visible = True
        txtArray(7).Visible = True
        txtArray(7).Text = mullenTargetTemperature(current_unit%)
        chkSet(7).Visible = True
        chkDelay(7).Visible = True
    End If
    
    chkUseTempControl.value = IIf(useTemperatureControlForAuto(current_unit%), 1, 0)
'    lblSet.Visible = useTemperatureControlForAuto(current_unit%)
'    lblDelay.Visible = useTemperatureControlForAuto(current_unit%)
    For i = 0 To 7
        lblArray(i).Enabled = useTemperatureControlForAuto(current_unit%)
        txtArray(i).Enabled = useTemperatureControlForAuto(current_unit%)
        chkSet(i).Enabled = useTemperatureControlForAuto(current_unit%)
        chkDelay(i).Enabled = useTemperatureControlForAuto(current_unit%)
        chkSet(i).value = IIf(setTemperatureForAuto(current_unit%, i), 1, 0)
        chkDelay(i).value = IIf(delayTestForTemperature(current_unit%, i), 1, 0)
    Next
End Sub
