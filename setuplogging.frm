VERSION 5.00
Begin VB.Form setuplogging 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Result Logging"
   ClientHeight    =   1050
   ClientLeft      =   390
   ClientTop       =   1380
   ClientWidth     =   10845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Log File"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "setuplogging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
uselog = (Check1.Value = 1)
WPPS Curr_U$, "uselog", Str$(Check1.Value), IFile$
logpath = Check1.Caption
WPPS Curr_U$, "logpath", logpath, IFile$
Unload Me
End Sub

Private Sub Command2_Click()
fsel_name$ = Check1.Caption
fsel_title$ = "Set Logging File"
fsel_io = False ' file doesn't have to exist
fsel Me.hwnd
If fsel_return$ <> "" Then
    Check1.Caption = fsel_return$
    Check1.Value = 1
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Check1.Caption = logpath
If uselog Then Check1.Value = 1
End Sub
