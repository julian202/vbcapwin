VERSION 5.00
Begin VB.Form temperaturelog 
   BackColor       =   &H000000FF&
   Caption         =   "Temperature Log"
   ClientHeight    =   2880
   ClientLeft      =   10065
   ClientTop       =   11985
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   3870
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   "Mullen Chamber"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hydro Head Chamber"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Reservoir"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Air"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cabinet"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bubbler"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Dry Chamber"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Wet Chamber"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   "10"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Log every:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2880
         Top             =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Minutes"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
   End
End
Attribute VB_Name = "temperaturelog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim channel(7) As Integer
Dim startTime As Single

Private Sub Check2_Click()
startTime = Timer - val(Text1.Text) * 60
End Sub

Private Sub Form_Load()
Dim i As Integer
channel(0) = reservoirTemperature
channel(1) = airTemperature
channel(2) = cabinetTemperature
channel(3) = bubblerTemperature
channel(4) = dryChamberTemperature
channel(5) = wetChamberTemperature
channel(6) = hydroHeadTemperature
channel(7) = mullenTemperature
For i = 0 To 7
    If channel(i) = 0 Then Check1(i).Enabled = False
Next i
'edc 12-11-06 alter border color and caption
'Me.Caption = Me.Caption & "    " & SubCaption
Me.BackColor = lngBorderColor
End Sub

Private Sub Form_Unload(cancel As Integer)
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Dim s(7) As Single
Dim i As Integer
Dim j As Integer
Dim lt As Single
For i = 0 To 7
    If Check1(i).value = 1 Then
        s(i) = readNewTemperature(channel(i))
        Label1(i) = str$(s(i))
    End If
Next i
If Check2.value = 0 Then Exit Sub
lt = Timer
If (lt < startTime) Then startTime = startTime - 86400
If (lt - startTime) >= val(Text1.Text) * 60 Then
    startTime = startTime + val(Text1.Text) * 60
    j = FreeFile
    Open EXE_Path + "temperaturelog.txt" For Append As #j
    Print #j, date$; ","; time$;
    For i = 0 To 5
        If Check1(i).value = 1 Then
            Print #j, ","; str$(s(i));
        End If
    Next i
    Print #j,
    Close #j
End If
End Sub
