VERSION 5.00
Begin VB.Form Msgform 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CapWin"
   ClientHeight    =   2580
   ClientLeft      =   2640
   ClientTop       =   2640
   ClientWidth     =   5550
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
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2580
   ScaleWidth      =   5550
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton Cancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton OK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.PictureBox lb 
         Height          =   135
         Left            =   1080
         ScaleHeight     =   75
         ScaleWidth      =   2775
         TabIndex        =   3
         Top             =   1560
         Width           =   2835
      End
      Begin VB.CommandButton StopTest 
         Appearance      =   0  'Flat
         Caption         =   "&Stop Test"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   3435
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Msgform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
DefSng A-Z
Dim ts$(1)                  ' Text strings for this form

Private Sub cancel_Click()
    Cancel_Aborted = True
    Unload Me
End Sub

Private Sub Form_Load()

    LoadTextStrings
    Me.top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(1) + str$(unitnumber)  ' "Unit"
    End If
    Me.BackColor = lngBorderColor
    
End Sub

Private Sub ok_Click()
    ok.Visible = False
End Sub

Private Sub StopTest_Click()
    Aborted = True
    StopTest.Visible = False
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    set_fontstuff Label1, default_font
    set_fontstuff Label, default_font
    ok.Caption = gpps2("msgform", "ok", language$, ok.Caption)
    set_fontname ok, default_font
    StopTest.Caption = gpps2("msgform", "stoptest", language$, StopTest.Caption)
    set_fontname StopTest, default_font
    
    ' Other text
    ts$(1) = gpps2("msgform", "ts1", language$, "Unit")

End Sub

