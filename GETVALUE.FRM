VERSION 5.00
Begin VB.Form GetValue 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Value"
   ClientHeight    =   2085
   ClientLeft      =   2700
   ClientTop       =   2700
   ClientWidth     =   5550
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
   ScaleHeight     =   2085
   ScaleWidth      =   5550
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   5055
      End
      Begin VB.CommandButton Cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Continue 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Check this box to allow users to modify this parameter at runtime."
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5055
      End
   End
End
Attribute VB_Name = "GetValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim ts$(1)              ' Text strings for this form

Private Sub cancel_Click()
    Got_Value = -9
    Got_Value_Check = Check1.value
    If elev_lqperm_ck = True Then elev_lqperm_ck = False
    Unload Me
End Sub

Private Sub Continue_Click()

    On Error Resume Next
    
    Got_Value = myVal(Text1.Text)
    ' trap out overflow
    If Err.Number = 6 Then Got_Value = 9E+99
    On Error GoTo 0
    Got_Text = Text1.Text
    Got_Value_Check = Check1.value
    Unload Me

End Sub

Private Sub Form_Load()

    LoadTextStrings
    'RemoveSysMenu Me
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(1) + Str$(unitnumber) ' "Unit"
    End If
    If supervisor = True And TType%(current_unit) = 4 And TMode(current_unit) = 2 And elev_lqperm_ck = True Then
       Check1.Visible = True
    Else
        Check1.Visible = False
    End If
    
    
    ' Superseded by LoadTextStrings:
    'Text1.FontName = system_font
    'Text1.fontsize = font_size
    'Text1.fontbold = font_bold
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor

End Sub

Private Sub Text1_Change()
    
    If Text1.Text = "" Then
        cancel.default = True
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    If Label1.Tag = "text" Then
        Continue.default = True
    Else
        Select Case KeyAscii
            Case 8, Asc(Decimal_Point$), 48 To 57
                Continue.default = True
            Case Else
                KeyAscii = 0
        End Select
    End If

End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    GetValue.Caption = gpps2("getvalue", "window title", language$, GetValue.Caption)
    Continue.Caption = gpps2("getvalue", "continue", language$, Continue.Caption)
    set_fontname Continue, default_font
    cancel.Caption = gpps2("getvalue", "cancel", language$, cancel.Caption)
    set_fontname cancel, default_font
    set_fontstuff Label1, default_font
    set_fontstuff Text1, default_font
    
    ' Other text
    ts$(1) = gpps2("getvalue", "ts1", language$, "Unit")

End Sub
