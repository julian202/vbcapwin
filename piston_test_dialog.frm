VERSION 5.00
Begin VB.Form piston_test_dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Piston Tester"
   ClientHeight    =   3690
   ClientLeft      =   12360
   ClientTop       =   7935
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox piston_travel_time_text 
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   3360
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   3360
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox num_cycles_text 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox pressure_text 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton StopButton 
      Caption         =   "Stop"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton StartButton 
      Caption         =   "Start"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Piston Travel Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "5. Wait for process to run the requested number of cycles or press 'Stop'"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   3120
      Width           =   5295
   End
   Begin VB.Label Label7 
      Caption         =   "4. Press 'Start' button."
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label Label6 
      Caption         =   "3. Make sure door is closed."
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "2. Enter number of cycles."
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "1. Enter required compression pressure."
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Piston Test Procedure:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Piston Cycles:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Compression Pressure:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "piston_test_dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim stopPistonTest As Boolean
Dim num_cycles As Integer
Dim piston_wait_time As Single

Private Sub ExitButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Zero_Reg
    close_v2_completely
    Move_Valve 0, "C"
    Me.pressure_text.Text = ""
    Me.num_cycles_text.Text = ""
    Me.StartButton.Enabled = True
    Me.StopButton.Enabled = False
End Sub

Private Sub StartButton_Click()
    If pressure_text.Text = "" Or val(pressure_text.Text) < 10 Then
        MsgBox "Please enter a pressure greater than 10"
        Exit Sub
    End If
    
    If num_cycles_text.Text = "" Or val(num_cycles_text.Text) < 1 Then
        MsgBox "Please enter a number of test cycles greater than 1"
        Exit Sub
    End If
    
    If piston_travel_time_text.Text = "" Or val(piston_travel_time_text) < 1 Then
        MsgBox "Please enter a travel time greater than 1 minute"
        Exit Sub
    End If
    
    stopPistonTest = False
    StartButton.Enabled = False
    StopButton.Enabled = True
    
    num_cycles = CInt(num_cycles_text.Text)
    num_cycles_text.Text = ""
    
    move_compression_regulator_to_pressure (val(pressure_text.Text))
    pressure_text.Text = ""
    
    piston_wait_time = (val(piston_travel_time_text) * 60)
    test_piston
End Sub

Private Sub StopButton_Click()
    stopPistonTest = True
    StartButton.Enabled = True
    StopButton.Enabled = False
End Sub

Private Sub test_piston()
    Dim i As Integer
    Picture1.Picture = LoadPicture(App.path & "\arrow up.bmp")
    Picture2.Picture = LoadPicture(App.path & "\arrow down.bmp")
    
    move_piston "C"
    waitseconds 10
    
    For i = 1 To num_cycles
        check_safety_door False
        
        Me.Caption = "Piston Tester - Cycle #" + Str$(i) + " of " + Str$(num_cycles)
        move_piston "O"
        
        Picture1.ZOrder 1
        Picture2.ZOrder 0
        Refresh
        waitseconds piston_wait_time
        
        check_safety_door False
        move_piston "C"
        
        Picture1.ZOrder 0
        Picture2.ZOrder 1
        Refresh
        
        waitseconds piston_wait_time
        
        If stopPistonTest Then
            i = num_cycles
        End If
    Next i
    
    StartButton.Enabled = True
    StopButton.Enabled = False
End Sub
