VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFastTest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fast Test"
   ClientHeight    =   7875
   ClientLeft      =   9540
   ClientTop       =   1530
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFastTest 
      Caption         =   "End Test"
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdFastTest 
      Caption         =   "Stop Test"
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   10
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdFastTest 
      Caption         =   "Start Test"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   1455
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Faster"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   8
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "More Accurate"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Stability"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Wet/Dry Test"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Bubble Point Test"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmFastTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

