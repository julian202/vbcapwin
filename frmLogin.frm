VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capwin Login"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ComboBox cmbUsers 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1710
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   4155
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblUsername 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogin_Click()
    Dim Index As Integer
    Dim userPassword As String
    Dim typedPassword As String
    
    Index = cmbUsers.ListIndex + 1
    userPassword = UAC_Users(Index).Password
    typedPassword = txtPassword.Text
    If typedPassword = userPassword Then
        UAC_CurrentUser = Index
        If UAC_Users(UAC_CurrentUser).accessLevel = 0 Then
            supervisor = True
        Else
            supervisor = False
        End If
        WPPS "UAC", "UAC_CurrentUser", Trim$(str$(UAC_CurrentUser)), UACFile
        UAC_LoggedIn = True
        Unload Me
    Else
        MsgBox "Incorrect password.", vbInformation
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 1 To UAC_UserCount
        cmbUsers.AddItem UAC_Users(i).Username
    Next i
    cmbUsers.ListIndex = UAC_CurrentUser - 1
End Sub

Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdLogin_Click
    End If
End Sub
