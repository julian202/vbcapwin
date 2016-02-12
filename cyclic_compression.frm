VERSION 5.00
Begin VB.Form cyclic_compression_setup 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cyclic Compression Setup"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "cyclic_compression.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Start"
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Compression Pressure:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Compression Time:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Decompression Time:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of Cycles:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "PSI"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Seconds"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Seconds"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Make sure sample is loaded and ready for testing before starting"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   6615
      End
   End
End
Attribute VB_Name = "cyclic_compression_setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ts$(6)                  ' Text strings for this form

Private Sub Command1_Click()

    Dim temp_pressure As Single, temp_down As Single, temp_up As Single, temp_num As Single

temp_pressure = myVal(Text1(0).Text) / PCNV
temp_down = myVal(Text1(1).Text)
temp_up = myVal(Text1(2).Text)
temp_num = myVal(Text1(3).Text)
If autocompress Then ' only test pressure if not autopiston
  If temp_pressure <= 0 Then
    MsgBox ts$(1)               ' "Error: Pressure must be above 0"
    Exit Sub
  End If
  If creg_table_pres!(creg_table_size%) < temp_pressure Then
    MsgBox ts$(2)               ' ("Compression Pressure greater than maximum calibrated pressure for regulator.")
    Text1(0).Text = Format$(creg_table_pres!(creg_table_size%) * PCNV)
    Exit Sub
  End If
End If
If temp_down <= 0 Then
    MsgBox ts$(3)               ' "Compression Time must be positive"
    Exit Sub
End If
If temp_up <= 0 Then
    MsgBox ts$(4)               ' "Decompression Time must be positive"
    Exit Sub
End If
If temp_num < 1 Then
    MsgBox ts$(5)               ' "Number of Cycles must be at least 1"
    Exit Sub
End If
If temp_num > 32767 Then
    MsgBox ts$(6)               ' "Number of Cycles is too large"
    Exit Sub
End If
cyclic_compression_pressure = temp_pressure
cyclic_compression_timedown = temp_down
cyclic_compression_timeup = temp_up
cyclic_compression_numcycles = temp_num
save_user_global_stuff
Aborted = False
Unload Me

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    LoadTextStrings
    Text1(0).Text = Format$(cyclic_compression_pressure * PCNV)
    Text1(1).Text = Format$(cyclic_compression_timedown)
    Text1(2).Text = Format$(cyclic_compression_timeup)
    Text1(3).Text = Format$(cyclic_compression_numcycles)
    Label2(0).Caption = PU$
    Text1(0).Enabled = autocompress ' can't do a pressure if you only have the autopiston
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    cyclic_compression_setup.Caption = gpps2("ccs", "window title", language$, cyclic_compression_setup.Caption)
    For i = 0 To 3
        Label1(i) = get_thing("ccsetup", "label1" + Str$(i), language$, Label1(i).Caption, Label1(i), default_font)
        Label2(i) = get_thing("ccsetup", "label2" + Str$(i), language$, Label2(i).Caption, Label2(i), default_font)
    Next i
    Command1.Caption = gpps2("ccsetup", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    Command2.Caption = gpps2("ccsetup", "command2", language$, Command2.Caption)
    set_fontname Command2, default_font
    
    ' Other text
    ts$(1) = gpps2("ccsetup", "ts1", language$, "Error: Pressure must be above 0")
    ts$(2) = gpps2("ccsetup", "ts2", language$, "Compression pressure greater than maximum calibrated pressure for regulator.")
    ts$(3) = gpps2("ccsetup", "ts3", language$, "Compression time must be positive")
    ts$(4) = gpps2("ccsetup", "ts4", language$, "Decompression time must be positive")
    ts$(5) = gpps2("ccsetup", "ts5", language$, "Number of cycles must be at least 1")
    ts$(6) = gpps2("ccsetup", "ts6", language$, "Number of cycles is too large")
    
End Sub

