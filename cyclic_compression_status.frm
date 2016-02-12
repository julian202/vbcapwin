VERSION 5.00
Begin VB.Form cyclic_compression_status 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cyclic Compression Status"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "cyclic_compression_status.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "&Abort"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Pressure: "
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Cycle Number:   of"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Compression Seconds Remaining:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   4455
      End
   End
End
Attribute VB_Name = "cyclic_compression_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Aborted = True
    Command1.Enabled = False
End Sub

Private Sub Form_Load()
    LoadTextStrings
    'edc 12-11-06 alter border color and caption
    'Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    cyclic_compression_status.Caption = gpps2("ccstatus", "window title", language$, cyclic_compression_status.Caption)
    Label1.Caption = get_thing("ccstatus", "label1", language$, Label1.Caption, Label1, default_font)
    Label2.Caption = get_thing("ccstatus", "label2", language$, Label2.Caption, Label2, default_font)
    Label3.Caption = get_thing("ccstatus", "label3", language$, Label3.Caption, Label3, default_font)
    Command1.Caption = gpps2("ccstatus", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font

End Sub

