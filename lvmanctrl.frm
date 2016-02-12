VERSION 5.00
Begin VB.Form lv_man_ctrl 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquid Vapor Manual Control"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   ControlBox      =   0   'False
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
   Icon            =   "lvmanctrl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6630
   ScaleWidth      =   7605
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.CheckBox Check1 
         Caption         =   "+2v"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pressure1"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Quit 
         Caption         =   "&Quit"
         Height          =   375
         Left            =   1200
         TabIndex        =   20
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pressure2"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton selfile 
         Caption         =   "Select File"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5400
         TabIndex        =   37
         Top             =   4320
         Width           =   1815
         Begin VB.OptionButton Option5 
            Caption         =   "H"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            Caption         =   "M"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   18
            Top             =   0
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            Caption         =   "S"
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   19
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.TextBox timeinc 
         Height          =   285
         Left            =   5400
         TabIndex        =   16
         Top             =   3960
         Width           =   1815
      End
      Begin VB.CommandButton startlog 
         Caption         =   "Start Logging"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton stoplog 
         Caption         =   "Stop Logging"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6120
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6120
         TabIndex        =   11
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pressurize"
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6000
         TabIndex        =   14
         Top             =   2880
         Width           =   1095
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1935
         LargeChange     =   100
         Left            =   240
         Max             =   1000
         Min             =   50
         SmallChange     =   50
         TabIndex        =   21
         Top             =   1920
         Value           =   500
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Run Auto Test"
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Gnd"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   6120
         TabIndex        =   15
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   6240
         TabIndex        =   27
         Top             =   5880
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "P1"
         Height          =   255
         Index           =   0
         Left            =   2820
         TabIndex        =   54
         Top             =   3375
         Width           =   375
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   2820
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label ValveClick 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   32
         Top             =   3720
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   4080
         Shape           =   3  'Circle
         Top             =   3720
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   6
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   3000
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   7
         Left            =   4080
         Shape           =   3  'Circle
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "P2"
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   53
         Top             =   1695
         Width           =   375
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   1
         Left            =   1680
         Top             =   1680
         Width           =   375
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   8
         Left            =   3120
         Shape           =   3  'Circle
         Top             =   3960
         Width           =   255
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   4440
         Top             =   3960
         Width           =   495
      End
      Begin VB.Line Line15 
         X1              =   3240
         X2              =   3240
         Y1              =   3840
         Y2              =   4320
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "pulse"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   64
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label ValveClick 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   33
         Top             =   3720
         Width           =   255
      End
      Begin VB.Line Line20 
         X1              =   3360
         X2              =   4680
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   0
         Left            =   1440
         Top             =   2880
         Width           =   735
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1440
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFC0&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   4440
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   57
         Top             =   3120
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   840
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label ValveClick 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   29
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   3615
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label ValveClick 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   28
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label reg_bottom_part 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   600
         TabIndex        =   55
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label ValveClick 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   30
         Top             =   2400
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   2280
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label ValveClick 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   31
         Top             =   2040
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   2280
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   52
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   8
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Line Line2 
         X1              =   3000
         X2              =   3000
         Y1              =   3720
         Y2              =   3840
      End
      Begin VB.Line Line3 
         X1              =   4680
         X2              =   4680
         Y1              =   3840
         Y2              =   3960
      End
      Begin VB.Line Line4 
         X1              =   2400
         X2              =   1800
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line5 
         X1              =   1800
         X2              =   1800
         Y1              =   3480
         Y2              =   3840
      End
      Begin VB.Line Line6 
         X1              =   1800
         X2              =   960
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line7 
         X1              =   960
         X2              =   960
         Y1              =   3360
         Y2              =   3720
      End
      Begin VB.Line Line8 
         X1              =   1800
         X2              =   1800
         Y1              =   2880
         Y2              =   2040
      End
      Begin VB.Line Line9 
         X1              =   1800
         X2              =   960
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line10 
         X1              =   960
         X2              =   960
         Y1              =   3120
         Y2              =   2640
      End
      Begin VB.Line Line11 
         X1              =   2280
         X2              =   1800
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line12 
         X1              =   1800
         X2              =   2280
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line13 
         X1              =   2520
         X2              =   3240
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line14 
         X1              =   2520
         X2              =   3240
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label1 
         Caption         =   "Vent"
         Height          =   255
         Left            =   3360
         TabIndex        =   51
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Vacuum"
         Height          =   255
         Left            =   3360
         TabIndex        =   50
         Top             =   2040
         Width           =   735
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   3720
         Width           =   255
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00FF80FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label ofilename 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   4800
         Width           =   6855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Time Increment"
         Height          =   255
         Left            =   5400
         TabIndex        =   49
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Line Line16 
         X1              =   3960
         X2              =   3960
         Y1              =   3480
         Y2              =   3840
      End
      Begin VB.Line Line17 
         X1              =   3840
         X2              =   3960
         Y1              =   3720
         Y2              =   3840
      End
      Begin VB.Line Line18 
         X1              =   4080
         X2              =   3960
         Y1              =   3720
         Y2              =   3840
      End
      Begin VB.Line Line19 
         X1              =   3360
         X2              =   3360
         Y1              =   3840
         Y2              =   3120
      End
      Begin VB.Label ValveClick 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   34
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label ValveClick 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   35
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Vacuum"
         Height          =   255
         Left            =   4800
         TabIndex        =   48
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Min. Pressure:"
         Height          =   255
         Left            =   4560
         TabIndex        =   47
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Max. Pressure:"
         Height          =   255
         Left            =   4560
         TabIndex        =   46
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Valve Pulse"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "500 ms"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label ValveClick 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   36
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Delta-Pressure:"
         Height          =   255
         Left            =   4680
         TabIndex        =   43
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Auto Test uses Delta, Max, Min,  Time Increment, and output file"
         Height          =   855
         Left            =   5520
         TabIndex        =   42
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label15 
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   5400
         Width           =   6855
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Pressure Mult.:"
         Height          =   255
         Left            =   4560
         TabIndex        =   40
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Min. DeltaP:"
         Height          =   255
         Left            =   4680
         TabIndex        =   39
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Pressure:"
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "I"
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   65
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "E"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   59
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "F"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   62
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "H"
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   61
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "G"
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   60
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "A"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   56
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "C"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   58
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "pulse"
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   63
         Top             =   3960
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   2640
         X2              =   4680
         Y1              =   3840
         Y2              =   3840
      End
   End
End
Attribute VB_Name = "lv_man_ctrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click(Index As Integer)
Label2(Index).Caption = ""
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Command3.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
pending$ = pending$ + Chr$(60)
End Sub

Private Sub Command2_Click()
abort_lv_goto = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Command1.Enabled = False
Command3.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
pending$ = pending$ + Chr$(61)
End Sub

Private Sub Form_Load()
Dim i As Integer
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
pending$ = ""
manual_data_logging = False
If lvperm_numvalves = 5 Then
    ' turn off valves 6 through 9
    Line16.Visible = False
    Line17.Visible = False
    Line18.Visible = False
    Line19.Visible = False
    Line20.Visible = False
    For i = 5 To 8
        Label3(i).Visible = False
        ValveClick(i).Visible = False
        ValveFill(i).Visible = False
    Next i
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Text1.Visible = False
    Text2.Visible = False
    Command1.Visible = False
    Command2.Visible = False
    Label11(0).Visible = False
    Label11(1).Visible = False
End If
lv_valve_pulse_timing = 0.5
'edc 12-11-06 alter border color and caption
Me.Caption = Me.Caption & "    " & SubCaption
Me.BackColor = lngBorderColor
End Sub

Private Sub Label11_Click(Index As Integer)
' pulse valves e&f (or g&h if Index=1)
' only do this if both valves are enabled
If ValveClick(4 + Index * 2).Enabled And ValveClick(5 + Index * 2).Enabled Then
  ValveClick(4 + Index * 2).Enabled = False
  ValveClick(5 + Index * 2).Enabled = False
  pending$ = pending$ + Chr$(40 + Index)
End If
End Sub

Private Sub Label3_Click(Index As Integer)
' when they click a letter, we pulse the valve
' you can only do this while the valveclick is enabled
If ValveClick(Index).Enabled Then
    ValveClick(Index).Enabled = False
    pending$ = pending$ + Chr$(Index + 20)
End If
End Sub

Private Sub Quit_Click()
Quit.Enabled = False
End Sub

Private Sub selfile_Click()
    fsel_title$ = "OUTPUT FILE"
    fsel_path$ = manual_data_path$
    fsel_name$ = ""
    fsel_io = False
    fsel Me.hwnd
    If fsel_return$ <> "" Then
        ofilename.Caption = fsel_return$
        manual_data_path$ = fsel_path$
        If Not manual_data_logging Then
            startlog.Enabled = True
        End If
    End If
End Sub

Private Sub startlog_Click()
manual_data_logging = True
startlog.Enabled = False
stoplog.Enabled = True
End Sub

Private Sub stoplog_Click()
manual_data_logging = False
startlog.Enabled = True
stoplog.Enabled = False
End Sub

Private Sub ValveClick_Click(Index As Integer)
ValveClick(Index).Enabled = False
pending$ = pending$ + Chr$(Index)
End Sub

Private Sub VScroll1_Change()
Dim i As Integer
Static internal_change As Boolean
If Not internal_change Then
    internal_change = True
    i = VScroll1.value / 50
    VScroll1.value = i * 50
    lv_valve_pulse_timing = VScroll1.value / 1000#
    Label10.Caption = Format$(VScroll1.value) + " ms"
    internal_change = False
End If
End Sub

Private Sub VScroll1_Scroll()
Label10.Caption = Format$(VScroll1.value) + " ms"
End Sub
