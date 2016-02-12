VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form BPPurgeDialog 
   Caption         =   "Machine Purge"
   ClientHeight    =   1920
   ClientLeft      =   6900
   ClientTop       =   2415
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   5895
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   345
      Left            =   4785
      TabIndex        =   2
      Top             =   1485
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   5685
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   270
         Left            =   570
         TabIndex        =   3
         Top             =   945
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   $"BPPurgeDialog.frx":0000
         Height          =   855
         Left            =   225
         TabIndex        =   1
         Top             =   240
         Width           =   5265
      End
   End
End
Attribute VB_Name = "BPPurgeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancel As Boolean

Private Sub Command1_Click()

Dim count%, i%, THIS&, POS&
If Command1.Caption = "Next" Then
    Command1.Caption = "Cancel"
    Label1.Caption = "Opening motor valve to 20%"
    Call Send_RS232("OB")
    Waitms 15000, True
'    Dim startTime&
'    startTime = Timer
'    Do: DoEvents
''        POS& = RSEcho("RH", 3)
''        THIS& = (POS& / ((oLimit - cLimit) + cLimit)) * 100
'        POS& = getMV2Position()
'        THIS& = POS& * 100
'        If cancel = True Then
'            Call Send_RS232("SB")
'            Exit Do
'        End If
'        If startTime - Timer > 10 Then Exit Do
'    Loop Until THIS& >= 10
    Call Send_RS232("SB")
    Label1.Caption = "Opening purge valve..."
    Call Move_Valve(2, "O")
    
    Label1.Caption = "Increasing regulator..."
    Call inc_reg(BPPostPurgeCounts)
    
    Label1.Caption = "Purging fluid for " + str$(BPPostPurgeDuration) + " Seconds..."
    For i% = 1 To BPPostPurgeDuration
        If cancel = True Then Exit For
        Waitms 1000, False
        ProgressBar1.value = ProgressBar1.value + 1
    Next i%
    Label1.Caption = "Zeroing regulator..."
    Call Zero_Reg
    Waitms 2000, False
    
    Label1.Caption = "Closing purge valve..."
    Call Move_Valve(2, "C")
    
    Label1.Caption = "Closing motor valve..."
    Send_RS232l "G-" + mv1_index_char, cLimit
    While RSEcho("V" + mv1_index_char, 1) <> Asc("S")
        If geoPoreValve = True Then
            Send_RS232l "G-" + mv1_index_char, V2POS
        End If
    Wend
    Label1.Caption = "Purge Complete!"
    Command1.Caption = "Close"
    Exit Sub
End If
If Command1.Caption = "Cancel" Then
    cancel = True
    Label1.Caption = "Canceling, please wait..."
    Waitms 2000, False
    Exit Sub
End If

If Command1.Caption = "Close" Then
    Unload Me
End If


End Sub


Private Sub Command2_Click()
Dim POS&
Dim THIS&
ReadXReturnX4 3

POS& = x5
THIS& = (POS& / ((oLimit - cLimit) + cLimit)) * 100
MsgBox THIS&
End Sub


Private Sub Command3_Click()

MsgBox Format$(CRAP * 100, "###.##")
End Sub


Private Sub Form_Load()

ProgressBar1.max = BPPostPurgeDuration
End Sub


