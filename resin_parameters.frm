VERSION 5.00
Begin VB.Form resin_parameters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resin Parameters"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox drain_time_text 
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox delay_text 
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox points_text 
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox increment_text 
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox start_pressure_text 
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox start_height_text 
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox fill_text 
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cleanout Drain Time (seconds):"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Stability Delay (seconds):"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Data Points:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Increment Pressure (psi):"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Pressure (psi):"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Height (cm):"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fill Height (cm):"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "resin_parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Resin_Fill_Height = val(fill_text.Text)
Resin_Start_Height = val(start_height_text.Text)
Resin_Start_Pressure = val(start_pressure_text.Text)
Resin_Increment_Pressure = val(increment_text.Text)
Resin_Number_Points = val(points_text.Text)
Resin_Stable_Seconds = val(delay_text.Text)
Resin_Drain_Seconds = val(drain_time_text.Text)

WPPS "capstuff", "Resin_Fill_Height", Resin_Fill_Height, CSFile$
WPPS "capstuff", "Resin_Start_Height", Resin_Start_Height, CSFile$
WPPS "capstuff", "Resin_Drain_Seconds", Resin_Drain_Seconds, CSFile$
WPPS "capstuff", "Resin_Start_Pressure", Resin_Start_Pressure, CSFile$
WPPS "capstuff", "Resin_Increment_Pressure", Resin_Increment_Pressure, CSFile$
WPPS "capstuff", "Resin_Number_Points", Resin_Number_Points, CSFile$
WPPS "capstuff", "Resin_Stable_Seconds", Resin_Stable_Seconds, CSFile$

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
fill_text.Text = str$(Resin_Fill_Height)
start_height_text.Text = str$(Resin_Start_Height)
start_pressure_text.Text = str$(Resin_Start_Pressure)
increment_text.Text = str$(Resin_Increment_Pressure)
points_text.Text = str$(Resin_Number_Points)
delay_text.Text = str$(Resin_Stable_Seconds)
drain_time_text.Text = str$(Resin_Drain_Seconds)
End Sub

