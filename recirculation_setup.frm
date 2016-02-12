VERSION 5.00
Begin VB.Form recirculation_setup 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recirculation System Setup"
   ClientHeight    =   5205
   ClientLeft      =   2385
   ClientTop       =   1875
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8550
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton Command1 
         Caption         =   "Open Chamber"
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close Chamber"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Stop All Flow"
         Height          =   615
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Partial Recirculation"
         Height          =   615
         Index           =   3
         Left            =   3360
         TabIndex        =   7
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Full Recirculation"
         Height          =   615
         Index           =   4
         Left            =   5760
         TabIndex        =   8
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Continue with Test Setup"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   4320
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel Test Setup"
         Height          =   375
         Index           =   6
         Left            =   4680
         TabIndex        =   10
         Top             =   4320
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open Door"
         Height          =   615
         Index           =   7
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Drain Chamber"
         Height          =   375
         Index           =   8
         Left            =   3840
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close Drain"
         Height          =   375
         Index           =   9
         Left            =   6120
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Chamber Status: Unknown - assumed open"
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Flow Status: Unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   6615
      End
      Begin VB.Label Label3 
         Caption         =   "Penetrometer Status: "
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   7695
      End
      Begin VB.Label Label4 
         Caption         =   "Recirculating Fluid:"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3600
         Width           =   7695
      End
      Begin VB.Label Label5 
         Caption         =   "Fluid At Sample:"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   7695
      End
   End
End
Attribute VB_Name = "recirculation_setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ts$(5)              ' Text strings for this form

Private Sub Command1_Click(Index As Integer)
' before going into recirc2, always close the drain valve, if it exists

    If Index = 4 And valve_23_exists And Vpos(23) > 0 Then pending$ = pending$ + Chr$(10)
    pending$ = pending$ + Chr$(Index + 1)

End Sub

Private Sub Form_Load()

    LoadTextStrings
    
    Command1(7).Visible = doorlock
    If piston_status = 1 Then
        Label1.Caption = ts$(1)         ' "Chamber Status: Open"
        Command1(4).Enabled = False     ' can't go into recirc2 if chamber open
    ElseIf piston_status = 2 Then
        Label1.Caption = ts$(2)         ' "Chamber Status: Closed"
    Else
        piston_status = 0
        If compression_pressure <> 0 Then
            Command1(4).Enabled = False ' can't go into recirc2 if chamber open and we are compressing
        End If
    End If
    
    If valve_23_exists Then
        If Vpos(23) > 0 Then
            Command1(8).Enabled = False
            Command1(9).Enabled = True
        Else
            Command1(8).Enabled = True
            Command1(9).Enabled = False
        End If
    Else
        Command1(8).Visible = False
        Command1(9).Visible = False
    End If
    
    If flow_status = 1 Then
        Label2.Caption = ts$(3)         ' "Flow Status: All Stopped"
        Command1(4).Enabled = False     ' can't go directly into recirc2
    ElseIf flow_status = 2 Then
        Label2.Caption = ts$(4)         ' "Flow Status: Partial Recirculation"
    ElseIf flow_status = 3 Then
        Label2.Caption = ts$(5)         ' "Flow Status: Full Recirculation"
        Command1(0).Enabled = False     ' can't open chamber when in recirc2
        Command1(7).Enabled = False     ' can't open door when in recirc2
    Else
        flow_status = 0
        Command1(4).Enabled = False     ' can't go directly into recirc2
    End If
    
    'JF 3-4-10
    If Not compression And recirculation Then
        Command1(0).Visible = False
        Command1(1).Visible = False
    End If
    
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
End Sub
Public Sub LoadTextStrings()
' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    recirculation_setup.Caption = get_thing("recirc_setup", "window title", language$, recirculation_setup.Caption, recirculation_setup, default_font)
    Label1.Caption = get_thing("recirc_setup", "label1", language$, Label1.Caption, Label1, default_font)
    ' Label 2 has special font size/bold
    set_fontname Label2, default_font
    Label2.Caption = gpps2("recirc_setup", "label2", language$, Label2.Caption)
    Label3.Caption = get_thing("recirc_setup", "label3", language$, Label3.Caption, Label3, default_font)
    Label4.Caption = get_thing("recirc_setup", "label4", language$, Label4.Caption, Label4, default_font)
    Label5.Caption = get_thing("recirc_setup", "label5", language$, Label5.Caption, Label5, default_font)
    For i = 0 To 9
        set_fontname Command1(i), default_font
        Command1(i).Caption = gpps2("recirc_setup", "command1(" + Format$(i) + ")", language$, Command1(i).Caption)
    Next i
    
    ' Other text
    ts$(1) = gpps2("recirc_setup", "ts1", language$, "Chamber Status: Open")
    ts$(2) = gpps2("recirc_setup", "ts2", language$, "Chamber Status: Closed")
    ts$(3) = gpps2("recirc_setup", "ts3", language$, "Flow Status: All Stopped")
    ts$(4) = gpps2("recirc_setup", "ts4", language$, "Flow Status: Partial Recirculation")
    ts$(5) = gpps2("recirc_setup", "ts5", language$, "Flow Status: Full Recirculation")
    
End Sub

