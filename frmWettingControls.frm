VERSION 5.00
Begin VB.Form frmWettingControls 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wetting Controls"
   ClientHeight    =   3195
   ClientLeft      =   3330
   ClientTop       =   4215
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4215
   Begin VB.Frame frameMain 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox chkChamber 
         Caption         =   "Chamber 3"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkChamber 
         Caption         =   "Chamber 2"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkChamber 
         Caption         =   "Chamber 1"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdReverse 
         Caption         =   "Reverse"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   "Forward"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   1095
      End
      Begin VB.HScrollBar hsSpeed 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   1
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label lblSpeed 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   1920
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmWettingControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim status As String

Private Sub chkChamber_Click(Index As Integer)
    
    If chkChamber(Index).value = vbChecked Then
        Move_Valve (23 - Index), "O"
    Else
        Move_Valve (23 - Index), "C"
    End If
    
End Sub

Private Sub cmdForward_Click()
    Send_RS232b "sA", hsSpeed.value
    Send_RS232 "MAF"
    status = "F"
End Sub

Private Sub cmdReverse_Click()
    Send_RS232b "sA", hsSpeed.value
    Send_RS232 "MAR"
    status = "R"
End Sub

Private Sub cmdStop_Click()
    If status <> "S" Then
        Send_RS232b "sA", 0
        Send_RS232 "MA" & status
    End If
    status = "S"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Me.BackColor = lngBorderColor
    For i = 1 To number_of_wetting_valves
        chkChamber(i - 1).Enabled = True
    Next i
    hsSpeed.value = 128
    lblSpeed.Caption = "Speed: 128"
    
    For i = 1 To number_of_wetting_valves
        Move_Valve (i - 1), "C"
    Next i
    
    'optChamber(0).value = True
End Sub

Private Sub Form_Unload(cancel As Integer)
    Dim i As Integer
    
    If status <> "S" Then
        Send_RS232b "sA", 0
        Send_RS232 "MA: & status"
        For i = 1 To number_of_wetting_valves
            Move_Valve (i - 1), "C"
        Next i
    End If
End Sub

Private Sub hsSpeed_Change()
    If status <> "S" Then
        Send_RS232b "sA", hsSpeed.value
        Send_RS232 "MA" & status
        lblSpeed.Caption = "Speed: " & hsSpeed.value
    End If
End Sub

'Private Sub optChamber_Click(Index As Integer)
'    Dim i As Integer
'
'    For i = 1 To number_of_wetting_valves
'        Move_Valve (i - 1), "C"
'    Next i
'
'    If status <> "S" Then
'        Move_Valve (Index), "O"
'    End If
'End Sub
