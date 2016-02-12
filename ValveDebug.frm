VERSION 5.00
Begin VB.Form ValveDebug 
   Caption         =   "Form1"
   ClientHeight    =   9090
   ClientLeft      =   1890
   ClientTop       =   2325
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   465
      Left            =   1995
      TabIndex        =   2
      Top             =   8520
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   465
      Left            =   4230
      TabIndex        =   1
      Top             =   8490
      Width           =   2130
   End
   Begin VB.TextBox Text1 
      Height          =   8100
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   270
      Width           =   6150
   End
End
Attribute VB_Name = "ValveDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Command2_Click()
Text1.Text = vbNullString
End Sub


