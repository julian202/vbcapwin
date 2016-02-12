VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AirResistTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Air Resistivity Test"
   ClientHeight    =   3495
   ClientLeft      =   8055
   ClientTop       =   4275
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7920
   Begin VB.CommandButton Command7 
      Caption         =   "zero reg"
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "open v2"
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "inc reg"
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "dec reg"
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "stop v2"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "close v2"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
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
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   2115
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   3390
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2865
         TabIndex        =   19
         Text            =   "9.8"
         Top             =   1215
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   180
         Top             =   2205
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "txt"
         DialogTitle     =   "Select a save file"
         Filter          =   "Text Files|*.txt"
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   1740
         Width           =   4575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Select File"
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
         Left            =   600
         TabIndex        =   13
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   750
         Left            =   675
         Top             =   2220
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2850
         TabIndex        =   3
         Text            =   "5.33"
         Top             =   465
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2850
         TabIndex        =   2
         Text            =   "2.8"
         Top             =   825
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Motor Valve Start Position(%):"
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
         TabIndex        =   20
         Top             =   1260
         Width           =   2640
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status: Ready..."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2940
         Width           =   7455
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   495
         Left            =   8160
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Pressure(Pa): N/A"
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
         Left            =   4335
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Velocity(cm/s): N/A"
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
         Left            =   4365
         TabIndex        =   7
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Flow(cc/m): N/A"
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
         Left            =   4335
         TabIndex        =   6
         Top             =   270
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Desired Velocity(cm/s):"
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
         TabIndex        =   5
         Top             =   510
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Sample Diameter(cm):"
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
         TabIndex        =   4
         Top             =   885
         Width           =   2055
      End
   End
End
Attribute VB_Name = "AirResistTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sampleArea As Double
Dim desiredFlow As Double
Dim psiToPascal As Double
Dim startPressure As Double
Dim incPress As Boolean
Dim decPress As Boolean
Dim zeroPress As Boolean
Dim openV As Boolean
Dim closeV As Boolean
Dim stopV As Boolean

Dim pArray() As Double
Dim vArray() As Double

Function GetInteropPressure(velocity As Double) As Double
'ported from c# to vb6
'AW 2016
' for(int i = 0; i <= distensionML.Count ;i++)
'{
'    if(burstVolume < Convert.ToDouble(distensionML[i]))
'    {
'        double bigPointML = Convert.ToDouble(distensionML[i]);
'        double bigPointCM = Convert.ToDouble(distensionCM[i]);
'        double smallPointML = Convert.ToDouble(distensionML[i - 1]);
'        double smallPointCM = Convert.ToDouble(distensionCM[i - 1]);
'        double mlDiff = bigPointML - smallPointML;
'        double cmDiff = bigPointCM - smallPointCM;
'        double myDiff = burstVolume - smallPointML;
'
'        double thisThing = myDiff / mlDiff;
'
'        double thisThingNow = cmDiff * thisThing;
'
'        double thisThingHere = smallPointCM + thisThingNow;
'        string format = thisThingHere.ToString("#.0000");
'        return Convert.ToDouble(format);
'    }
'}

Dim i%
For i% = 0 To UBound(vArray()) - 1
    If velocity < vArray(i%) Then
        bigPointV = vArray(i%)
        bigPointP = pArray(i%)
        smallPointV = vArray(i% - 1)
        smallPointP = pArray(i% - 1)
        vDiff = bigPointV - smallPointV
        pDiff = bigPointP - smallPointP
        myDiff = velocity - smallPointV
        
        thisThing = myDiff - vDiff
        
        thisThingNow = pDiff * thisThing
        thisThingHere = smallPointV + thisThingNow
    End If
Next i%

GetInteropPressure = thisThingHere
End Function


Public Sub Status(txt As String)
Label7.Caption = txt
End Sub


Private Sub Command1_Click()
Dim XX&
Dim pDiff As Double
Status "Closing motor valve and zeroing regulator..."
Zero_Reg
X = close_v2_completely()
ReDim pArray(0)
ReDim vArray(0)
Dim position As Single
Waitms 1000, False
position = getMV2Position() * 100
Waitms 1000, False
Status "Opening motor valve to starting position..."
Send_RS232 ("OB")
Do: DoEvents
    position = getMV2Position() * 100
Loop Until position >= CDbl(Text4.Text)
Send_RS232 ("SB")
Status "Motor valve set..."
sampleArea = 3.14159265358979 * ((CDbl(Text2.Text) / 2) * (CDbl(Text2.Text) / 2))

desiredFlow = (CDbl(Text1.Text) * sampleArea) * 60
psiToPascal = 6894.75729
Status "Opening valve 1 for testing..."
Call Move_Valve(0, "O")


ReadXReturnX4 2
startPressure = x5 * psiToPascal
'Timer1.Enabled = True
Dim velocity As Double
Dim Flow As Double
Dim Pressure As Double
'FUSE% = 1
Dim counting As Integer
Do: DoEvents
    If Flow < desiredFlow Then
        If desiredFlow - Flow > 1000 Then
            inc_reg 50
            counting = counting + 50
            Status "Increasing flow 50 counts..."
        ElseIf desiredFlow - Flow > 500 Then
            inc_reg 25
            counting = counting + 25
            Status "Increasing flow 25 counts..."
        ElseIf desiredFlow - Flow > 100 Then
            inc_reg 10
            counting = counting + 10
            Status "Increasing flow 10 counts..."
        Else
            inc_reg 1
            counting = counting + 1
            Status "Increasing flow 1 count..."
        End If
    End If
    If Flow > desiredFlow Then
        lower_reg 1
        counting = counting + 1
        Status "Decreasing flow 1 count..."
    End If
    If counting >= 10 Then
        Status "Waiting for stability..."
        Waitms 2000, False
        counting = 0
    End If
    Label6.Caption = CStr(counting)
    ReadXReturnX4 1
    Flow = x5
    Label4.Caption = "Flow(cc/m): " + Format$(CStr(Flow), "##0.000")
    ReadXReturnX4 2
    Pressure = x5 * psiToPascal
    ReDim Preserve pArray(UBound(pArray) + 1)
    pArray(UBound(pArray) - 1) = Pressure
    Label3.Caption = "Pressure(Pa): " + Format$(CStr(Pressure), "########0.000")
    velocity = ((Flow / sampleArea) / 60)
    ReDim Preserve vArray(UBound(vArray) + 1)
    vArray(UBound(vArray) - 1) = velocity
    Label5.Caption = "Velocity(cm/s): " + Format$(CStr(velocity), "##0.000")
Loop Until velocity >= CDbl(Text1.Text)
Status "Test complete"

If velocity > CDbl(Text1.Text) Then
    Pressure = GetInteropPressure(CDbl(Text1.Text))
End If
pDiff = Pressure - startPressure
Open Text3.Text For Output As #101
Print #101, "Porous Materials, Inc"
Print #101, "20 Dutchmill Rd"
Print #101, "Ithaca, NY 14850"
Print #101, "(607)257-5544"
Print #101, " "
Print #101, "Air Resistivity Summary"
Print #101, " "
Print #101, "Sample Diameter: " + Text2.Text; " cm"
Print #101, "Target Velocity: " + Text1.Text + " cm/s"
Print #101, "Actual Flow: " + Format(CStr(Flow), "#####0.0##") + " cc"
Print #101, "Pressure Change: "; Format$(CStr((pDiff)), "#######0.000") + " Pa"
Close #101

XX& = MsgBox("Test completed with a pressure change of " + Format$(CStr((pDiff)), "#######0.000") + " Pa. Would you like to view the summary now?", vbQuestion + vbYesNo, "Capwin")
If XX& = vbYes Then
    Call Shell("notepad " + Text3.Text, vbNormalFocus)
End If
Zero_Reg
'MsgBox CStr(startPressure) + " - " + CStr(Pressure) + " = " + CStr(Pressure - startPressure)

Unload Me

End Sub


Private Sub Command2_Click()
openV = True

End Sub

Private Sub Command3_Click()
closeV = True
End Sub


Private Sub Command4_Click()
stopV = True
End Sub


Private Sub Command5_Click()
incPress = True
End Sub

Private Sub Command6_Click()
decPress = True

End Sub

Private Sub Command7_Click()
zeroPress = True
End Sub


Private Sub Command8_Click()
CommonDialog1.ShowSave
Text3.Text = CommonDialog1.filename
Command1.Enabled = True
End Sub

Private Sub Form_Load()
Init_For_Ctrl (True)
Text1.Text = gpps2(Curr_U$, "AirResistVelocity", IFile$, "5.33")
Text2.Text = gpps2(Curr_U$, "AirResistDiameter", IFile$, "2.8")
Text4.Text = gpps2(Curr_U$, "AirResistMVStart", IFile$, "9.8")

End Sub


Private Sub Form_Unload(cancel As Integer)
WPPS Curr_U$, "AirResistVelocity", Text1.Text, IFile$
WPPS Curr_U$, "AirResistDiameter", Text2.Text, IFile$
WPPS Curr_U$, "AirResistMVStart", Text4.Text, IFile$
End Sub


Private Sub Timer1_Timer()
Dim velocity As Double
Dim Flow As Double
Dim Pressure As Double
FUSE% = 1
ReadXReturnX4 1
Flow = x5
Label4.Caption = "Flow(cc/m): " + CStr(Flow)
ReadXReturnX4 2
Pressure = x5
Label3.Caption = "Pressure(Pa): " + CStr(Pressure * psiToPascal)
velocity = ((Flow / sampleArea) / 60)
Label5.Caption = "Velocity(cm/s): " + CStr(velocity)

If openV Then
    openV = False
    Send_RS232 ("OB")
End If
If closeV Then
    closeV = False
    Send_RS232 ("CB")
End If
If stopV Then
    stopV = False
    Send_RS232 ("SB")
End If
If incPress Then
    incPress = False
    inc_reg 1
End If
If decPress Then
    decPress = False
    lower_reg 1
End If
If zeroPress Then
    zeroPress = False
    Zero_Reg
End If
End Sub


