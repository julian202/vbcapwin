VERSION 5.00
Begin VB.Form qcbubble 
   BackColor       =   &H000000FF&
   Caption         =   "QC Bubble Point Test"
   ClientHeight    =   3255
   ClientLeft      =   2055
   ClientTop       =   2265
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   8430
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "5"
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "6"
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "7"
         Height          =   255
         Index           =   6
         Left            =   5520
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "8"
         Height          =   255
         Index           =   7
         Left            =   6000
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "9"
         Height          =   255
         Index           =   8
         Left            =   6600
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "10"
         Height          =   255
         Index           =   9
         Left            =   7200
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Select Chambers:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "qcbubble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim i As Integer

LoadTextStrings
' this does not support manual multi chamber, but this whole form is not supported at this time anyway
If Not multiChamberSystem Then
    ' 1 chamber instruments do not have labels to select chambers
    Label1.Visible = False
    Check1(0).Visible = False
    Check1(0).value = 1
    Check1(1).Visible = False
    Check1(1).value = 0
End If
For i = chambers To 9
    Check1(i).Visible = False
    Check1(i).value = 0
Next i
For i = 0 To chambers - 1
    Check1(i).value = IIf(selchamber(i + 1), 1, 0)
Next i
'edc 12-11-06 alter border color and caption
Me.Caption = Me.Caption & "    " & SubCaption
Me.BackColor = lngBorderColor

End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    qcbubble.Caption = gpps2("qcbubble", "window title", language$, qcbubble.Caption)
    Label1.Caption = get_thing("qcbubble", "window title", language$, qcbubble.Caption, qcbubble, default_font)
    
End Sub
