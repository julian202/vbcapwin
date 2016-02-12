VERSION 5.00
Begin VB.Form status 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status"
   ClientHeight    =   2145
   ClientLeft      =   3705
   ClientTop       =   3720
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   6585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.Timer Timer1 
         Left            =   360
         Top             =   1200
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "&Abort"
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   960
         Width           =   6375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   6135
      End
   End
End
Attribute VB_Name = "status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ts$(1)              ' Text strings for this form

Private Sub Command1_Click()
    Command1.Enabled = False
End Sub

Private Sub Form_Load()

    LoadTextStrings
    Me.top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Command1.Enabled = True
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(1) + str$(unitnumber)            ' "Unit"
    End If
    'edc 12-11-06 alter border color and caption
    'Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    status.Caption = gpps2("status", "window title", language$, status.Caption)
    Command1.Caption = gpps2("status", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    set_fontstuff Label1, default_font
    set_fontstuff Label2, default_font
    
    ' Other text
    ts$(1) = gpps2("status", "ts1", language$, "Unit")
    
End Sub

