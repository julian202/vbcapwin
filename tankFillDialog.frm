VERSION 5.00
Begin VB.Form TankFillDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supply Tank Status"
   ClientHeight    =   2370
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1935
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Please press tank fill button to turn on pump"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Level:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "TankFillDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim exitLoop As Boolean
Private Sub Command1_Click()
    exitLoop = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    startTankMonitor
End Sub

Private Sub Form_Load()
    Command1.Enabled = True
    Command2.Enabled = True
    exitLoop = False
End Sub

Private Sub startTankMonitor()
    While x5 <= min_tank_fill_level And Not exitLoop
        ReadXReturnX4 tank_level_location
        Label2.Caption = Xformat((x5), "##0.0") + " %"
        waitseconds 0.5
    Wend
    Unload Me
End Sub
