VERSION 5.00
Begin VB.Form singlePointGasTest 
   BackColor       =   &H000000FF&
   Caption         =   "Single-Point Test"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4245
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label f 
         Alignment       =   1  'Right Justify
         Caption         =   "Flow (LPM)"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label f 
         Alignment       =   1  'Right Justify
         Caption         =   "Pressure"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label f 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of points recorded"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Label7"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Collecting data"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label f 
         Alignment       =   1  'Right Justify
         Caption         =   "Time elapsed (s)"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   2
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Label7"
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   3360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "singlePointGasTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 

Private Sub Form_Load()
    
    Dim i As Integer
    
    Me.Caption = gpps2("singlepoint", "window title", language$, Me.Caption)
    Label1.Caption = get_thing("singlepoint", "label1", language$, Label1.Caption, Label1, default_font)
    Label5.Caption = get_thing("singlepoint", "label5", language$, Label5.Caption, Label5, default_font)
    Label2.Caption = get_thing("singlepoint", "label2", language$, Label2.Caption, Label2, default_font)
    Label9.Caption = get_thing("singlepoint", "label9", language$, Label9.Caption, Label9, default_font)
    Label10.Caption = get_thing("singlepoint", "label10", language$, Label10.Caption, Label10, default_font)
    For i = 0 To 3
        f(i).Caption = get_thing("singlepoint", "f" + Format$(i), language$, f(i).Caption, f(i), default_font)
    Next i
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
End Sub
 
