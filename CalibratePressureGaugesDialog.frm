VERSION 5.00
Begin VB.Form CalibratePressureGaugesDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calibrate Pressure Gauges"
   ClientHeight    =   3195
   ClientLeft      =   7170
   ClientTop       =   2760
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CommandA 
      Caption         =   "Calibrate Pressures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton CommandShowPressures 
      Caption         =   "Show Pressures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton CommandOpenValve 
      Caption         =   "Open Valve To Low Pressure Gauge"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton CommandCloseValve 
      Caption         =   "Close Valve To Low Pressure Gauge"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "CalibratePressureGaugesDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandA_Click()
    Dialog.Label1.Caption = "Calibrating Pressure Gauges..."
    Dialog.Label6.Visible = False
    Dialog.Show
    TitleScrn.CalibratePressures
End Sub

Private Sub CommandCloseValve_Click()
 Move_Valve 10, "C" 'close valve to low pressure gaue
End Sub

Private Sub CommandOpenValve_Click()
    Move_Valve 10, "O" 'open valve to low pressure gaue
End Sub

Private Sub CommandShowPressures_Click()
    Dim hhp As String  ' hi range hi pressure gauge.
    Dim hlp As String  ' hi range lo pressure gauge.
    Dim hhpVal As Single
    Dim hlpVal As Single
    'Move_Valve 10, "O" 'open valve to low pressure gaue
    'Sleep 700 ' to sleep for 1 second for valve to low pressure gauge to open
    tempPres% = Pres%
    Pres% = 0 ' hi range hi pressure gauge.
    ReadXReturnX4 2
    hhp = Xformat$(x5 * PCNV, "#####.000")
    Pres% = 2 ' hi range low pressure gauge.
    ReadXReturnX4 2
    hlp = Xformat$(x5 * PCNV, "#####.000")
    Pres% = tempPres%
    Dialog.Label1.Caption = "Gauge Pressures"
    Dialog.Label6.Visible = False
    Dialog.Label2.Caption = hhp
    Dialog.Label3.Caption = hlp
    Dialog.Show
End Sub
