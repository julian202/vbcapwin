VERSION 5.00
Begin VB.Form frmFrazierPassFail 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frazier Pass/Fail"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox happyface 
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2040
      Picture         =   "frmFrazierPassFail.frx":0000
      ScaleHeight     =   495
      ScaleMode       =   0  'User
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   2040
      Width           =   500
   End
   Begin VB.PictureBox sadface 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2040
      Picture         =   "frmFrazierPassFail.frx":0D26
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   2040
      Width           =   500
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblTestFrazierValue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblMaxPassValue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblMinPassValue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblTestFrazierValueTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Test Frazier Value:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblMaxPassTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maximum Pass Value:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblMinPassTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Minimum Pass Value:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Test Passed!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmFrazierPassFail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub processData()
    If frazier_array.s > 0 Then
    
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub showPassed()
    setLabelColor vbGreen
    happyface.Visible = True
    sadface.Visible = False
End Sub

Private Sub showFailed()
    setLabelColor vbRed
    happyface.Visible = False
    sadface.Visible = True
End Sub

Private Sub Form_Load()
    lblMinPassValue.Caption = str$(minFrazierPass(current_unit%))
    lblMaxPassValue.Caption = str$(maxFrazierPass(current_unit%))
End Sub

Private Sub setLabelColor(newColor As Integer)
    lblStatus.ForeColor = newColor
    lblMinPassTitle.ForeColor = newColor
    lblMaxPassTitle.ForeColor = newColor
    lblTestFrazierValueTitle.ForeColor = newColor
    lblMinPassValue.ForeColor = newColor
    lblMaxPassValue.ForeColor = newColor
    lblTestFrazierValue.ForeColor = newColor
End Sub
