VERSION 5.00
Begin VB.Form lpParameters 
   Caption         =   "Liquid Permeametry Parameters"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStartingPressure 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtPointStepPressure 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtMaximumPressure 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtMaximumWaitBetweenPoints 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtMaximumNumberOfPoints 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Starting Pressure (PSI differential):"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Point Step Pressure (PSI):"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Maximum Pressure (PSI differential):"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Maximum Wait Between Points (sec):"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Maximum Number Of Points:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
End
Attribute VB_Name = "lpParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim startPressure As Double
Dim maxPressure As Double
Dim stepPressure As Double
Dim maxWait As Single
Dim maxPoints As Integer

Private Sub cmdCancel_Click()
    lperm_user_cancelled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    startPressure = txtStartingPressure.Text / PCNV
    If startPressure > max_liq_pres Then
        MsgBox "The maximum allowed starting pressure is " + str$(max_liq_pres * PCNV)
        Exit Sub
    End If
    If autocompress And Compression_Increase_Factor < 0 And compression_pressure <> 0 And startPressure * Abs(Compression_Increase_Factor) > compression_pressure Then
        MsgBox "The starting pressure is too high for the current compression pressure setting"
        Exit Sub
    End If
    
    stepPressure = txtPointStepPressure.Text / PCNV
    If stepPressure < 0 Then stepPressure = 0
    
    
    maxPressure = txtMaximumPressure.Text / PCNV
    If maxPressure > max_liq_pres Then
        MsgBox "The maximum allowed pressure is " + str$(max_liq_pres * PCNV)
        Exit Sub
    End If
    If maxPressure < startPressure Then
        MsgBox "The maximum pressure can not be less than the starting pressure"
        Exit Sub
    End If
    If autocompress And Compression_Increase_Factor < 0 And compression_pressure <> 0 And maxPressure * Abs(Compression_Increase_Factor) > compression_pressure Then
        MsgBox "Maximum pressure too high with current compression pressure setting"
        Exit Sub
    End If
    
    maxWait = txtMaximumWaitBetweenPoints.Text
    If maxWait = 0 Then maxWait = 9.999999E+35
    
    maxPoints = txtMaximumNumberOfPoints.Text
    If maxPoints = 0 Then maxPoints = 100
    
    lperm_startp = startPressure
    lperm_maxp = maxPressure
    lperm_stepp = stepPressure
    lperm_maxwait = maxWait
    lperm_maxpoints = maxPoints
    
    WPPS Curr_U$, "lperm_startp", str$(lperm_startp), IFile$
    WPPS Curr_U$, "lperm_maxp", str$(lperm_maxp), IFile$
    WPPS Curr_U$, "lperm_stepp", str$(lperm_stepp), IFile$
    WPPS Curr_U$, "lperm_maxwait", str$(lperm_maxwait), IFile$
    WPPS Curr_U$, "lperm_maxpoints", str$(lperm_maxpoints), IFile$
    
    Unload Me
End Sub

Private Sub Form_Load()
    startPressure = lperm_startp
    maxPressure = lperm_maxp
    stepPressure = lperm_stepp
    maxWait = lperm_maxwait
    maxPoints = lperm_maxpoints
    
    txtStartingPressure.Text = startPressure
    txtPointStepPressure.Text = stepPressure
    txtMaximumPressure.Text = maxPressure
    txtMaximumWaitBetweenPoints.Text = maxWait
    txtMaximumNumberOfPoints.Text = maxPoints
    
    lperm_user_cancelled = False
End Sub
