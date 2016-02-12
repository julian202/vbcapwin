VERSION 5.00
Begin VB.Form MVTest 
   Caption         =   "Geopore Valve Calibration"
   ClientHeight    =   4320
   ClientLeft      =   7875
   ClientTop       =   3000
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   10815
   Begin VB.TextBox Text1 
      Height          =   3885
      Left            =   4185
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   225
      Width           =   6450
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3300
      Top             =   3420
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Pulse Open"
      Height          =   585
      Left            =   435
      TabIndex        =   10
      Top             =   2370
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Auto"
      Height          =   600
      Left            =   1080
      TabIndex        =   9
      Top             =   3015
      Width           =   1995
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set Limits"
      Height          =   585
      Left            =   1545
      TabIndex        =   8
      Top             =   2370
      Width           =   1065
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pulse Close"
      Height          =   600
      Left            =   2670
      TabIndex        =   7
      Top             =   2355
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Interval        =   125
      Left            =   3300
      Top             =   3000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   600
      Left            =   1545
      TabIndex        =   6
      Top             =   1740
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   600
      Left            =   2655
      TabIndex        =   5
      Top             =   1725
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   600
      Left            =   435
      TabIndex        =   4
      Top             =   1740
      Width           =   1065
   End
   Begin VB.Label Label5 
      Caption         =   "Set Close Limit:"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   585
      TabIndex        =   12
      Top             =   765
      Width           =   2625
   End
   Begin VB.Label Label4 
      Caption         =   "Open %:"
      Height          =   390
      Left            =   885
      TabIndex        =   3
      Top             =   1320
      Width           =   2070
   End
   Begin VB.Label Label3 
      Caption         =   "Position:"
      Height          =   270
      Left            =   885
      TabIndex        =   2
      Top             =   1035
      Width           =   2625
   End
   Begin VB.Label Label2 
      Caption         =   "Close Limit:"
      Height          =   270
      Left            =   870
      TabIndex        =   1
      Top             =   495
      Width           =   2625
   End
   Begin VB.Label Label1 
      Caption         =   "Open Limit: "
      Height          =   270
      Left            =   855
      TabIndex        =   0
      Top             =   225
      Width           =   2625
   End
End
Attribute VB_Name = "MVTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub status(msg$)
Text1.Text = Text1.Text + msg$ + vbCrLf
End Sub

Private Sub Command1_Click()
openValve = True
End Sub

Private Sub Command2_Click()
closeValve = True
'closedPercent = CDbl(Text1.Text)

End Sub

Private Sub Command3_Click()
stopValve = True
End Sub

Private Sub Command4_Click()
Dim lastGeoValvePosition&, geoValvePosition&, geoOpenLimit&

Timer1.Enabled = False
status "turning off manual control"
Waitms 2000, False

MsgBox "Auto calibration of Geo Pore valve will begin..."

'open valve for 5 seconds.
status "opening valve"
Call Send_RS232("OB")
status "waiting 5 seconds"
Call Waitms(5000, False)
'stop valve
status "stopping valve"
Call Send_RS232("SB")
'save current valve position

lastGeoValvePosition& = getMV2Position()
lastGeoValvePosition& = x4
status "recording valve position: " + str$(lastGeoValvePosition&)
'start shutting valve
status "closing valve..."
Call Send_RS232("CB")
'loop
Do: DoEvents
    'wait .5 seconds
    status "waiting .5 seconds for counts to change"
    Call Waitms(500, False)
    'get valve position
    geoValvePosition& = getMV2Position()
    geoValvePosition& = x4
    status "recording valve position: " + str$(geoValvePosition&)
    'if position has not changed for .5 seconds exit loop
    If geoValvePosition& = lastGeoValvePosition& Then
        status "valve didn't move any more, last position: " + str$(lastGeoValvePosition&)
        status "current valve position: " + str(geoValvePosition&)
        Exit Do
    End If
    'if it has changed, save last position and continue loop
    
    lastGeoValvePosition& = geoValvePosition&
    status "valve is moving still, last position: " + str$(lastGeoValvePosition&)
Loop
'stop the valve
status "stopping valve"
Call Send_RS232("SB")
status "waiting 1 second"
Waitms 1000, False
'pulse closed
status "pulsing closed"
Call Send_RS232("DB")
'wait a second
status "waiting 2 seconds"
Call Waitms(2000, False)
'pulse closed
status "pulsing closed"
Call Send_RS232("DB")

'at this point valve should be closed.
'read position one more time, this will be our close limit
geoValvePosition& = getMV2Position()
geoValvePosition& = x4
'read open limit
geoOpenLimit& = RSEcho("RR", 3)
'record values to capwin.ini file
WPPS "Capstuff", "OLIMIT", str$(geoOpenLimit&), CSFile$
oLimit = geoOpenLimit&
WPPS "Capstuff", "CLIMIT", str$(geoValvePosition&), CSFile$
cLimit = geoValvePosition&
MsgBox "Auto calibration complete."

Waitms 2000, False
status "turning on manual control"
Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
pulseValveClosed = True

End Sub


Private Sub Command6_Click()
Dim POS&, openl&
'stop valve if it is moving
Call Send_RS232("SB")
'shut off timer
'Timer1.Enabled = False
'current position is close limit
POS& = RSEcho("RH", 3)
'open limit is the....open limit.
openl& = RSEcho("RR", 3)
'save stuffs
WPPS "Capstuff", "OLIMIT", str$(openl&), CSFile$
oLimit = openl&
WPPS "Capstuff", "CLIMIT", str$(POS&), CSFile$
cLimit = POS&
'say something to the user
MsgBox "Valve calibration complete!"
'go away
Unload Me
End Sub

Private Sub Command7_Click()
pulseValveOpen = True
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(cancel As Integer)
Timer1.Enabled = False
End Sub


Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Timer1_Timer()
Dim openl&, closeL&, POS As Double, percent&, span&
Dim THIS As Single

'do any pending commandses
If openValve = True Then
    openValve = False
    Call Send_RS232("OB")
End If
If closeValve = True Then
    closeValve = False
    closingValve = True
    Call Send_RS232("CB")
End If
If stopValve = True Then
    stopValve = False
    Call Send_RS232("SB")
    closingValve = False
End If
If pulseValveClosed = True Then
    pulseValveClosed = False
    Call Send_RS232("DB")
End If
If pulseValveOpen = True Then
    pulseValveOpen = False
    Call Send_RS232("IB")
End If
'read open/closed limit, position,

'get percent open
openl& = RSEcho("RR", 3)
'openL& = openL& - (DAC_span / 20)
Label1.Caption = "Open Limit: " + str$(openl&)
closeL& = RSEcho("RQ", 3)
Label2.Caption = "Close Limit: " + str$(closeL&)
Label5.Caption = "Set Close Limit: " + str$(cLimit)
THIS = getMV2Position()
POS = THIS * 100
Label3.Caption = "Position: " + str$(x4)
Label4.Caption = "Open: " + Format(POS, "#00.00")


End Sub


Private Sub Timer2_Timer()

valveTimer = valveTimer + 0.5

End Sub


