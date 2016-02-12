VERSION 5.00
Begin VB.Form stabilitydebug 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stability Debug Window"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Label Label1 
         Caption         =   "Stability Method Used:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Background Stability Method Information:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Time"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "MinP"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "MaxP"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "MinF"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "MaxF"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "info"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "info"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "info"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "info"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   4215
      End
   End
End
Attribute VB_Name = "stabilitydebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
stability_debug = True
'edc 12-11-06 alter border color and caption
'Me.Caption = Me.Caption & "    " & SubCaption
Me.BackColor = lngBorderColor
End Sub

Private Sub Form_Unload(cancel As Integer)
stability_debug = False
End Sub

