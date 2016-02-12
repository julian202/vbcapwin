VERSION 5.00
Begin VB.Form LiqSel 
   Caption         =   "Liquid Select"
   ClientHeight    =   2265
   ClientLeft      =   4440
   ClientTop       =   3090
   ClientWidth     =   4905
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
   LinkTopic       =   "GasSel"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2265
   ScaleWidth      =   4905
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2280
      TabIndex        =   6
      Text            =   " "
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2280
      TabIndex        =   4
      Text            =   " "
      Top             =   1200
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "X: Other Liquid"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Water"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Viscosity (cP):"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Name of Liquid:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Select Liquid:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "LiqSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ts$(4)                     ' Text strings for this form

Private Sub Command1_Click()

Dim a$, b$
Dim X As Single

If Option1(0).Value Then
    LiquidID$(current_unit%) = "W"
    Liquid$(current_unit%) = ts$(4)     ' "Water"
    Unload Me
Else
    If myVal(Text2.Text) <= 0 Then
        MsgBox ts$(1)           ' "Error: incorrect viscosity value"
        Exit Sub
    End If
    If Len(Text1.Text) = 0 Then
        MsgBox ts$(2)           ' "Error: incorrect liquid name value"
        Exit Sub
    End If
    a$ = Text1.Text
    X = myVal(Text2.Text)
    b$ = LTrim$(Str$(X))
    Liquid$(current_unit%) = Left$(b$ + Space$(20), 20) + a$
    Unload Me
End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

LoadTextStrings

If Option1(4).Value Then
    Text1.Visible = True
    Text2.Visible = True
    Label2.Visible = True
    Label3.Visible = True
Else
    Text1.Visible = False
    Text2.Visible = False
    Label2.Visible = False
    Label3.Visible = False
End If
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(3) + Str$(unitnumber)       ' "Unit"
    End If
    
    ' Superseded by LoadTextStrings
    'Text1.FontName = system_font
    'Text1.fontsize = font_size
    'Text1.fontbold = font_bold

End Sub

Private Sub Option1_Click(Index As Integer)

If Option1(4).Value Then
    Text1.Visible = True
    Text2.Visible = True
    Label2.Visible = True
    Label3.Visible = True
Else
    Text1.Visible = False
    Text2.Visible = False
    Label2.Visible = False
    Label3.Visible = False
End If

End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    LiqSel.Caption = gpps2("liqsel", "window title", language$, LiqSel.Caption)
    Label1.Caption = get_thing("liqsel", "label1", language$, Label1.Caption, Label1, default_font)
    Label2.Caption = get_thing("liqsel", "label2", language$, Label2.Caption, Label2, default_font)
    Label3.Caption = get_thing("liqsel", "label3", language$, Label3.Caption, Label3, default_font)
    Option1(0).Caption = get_thing("liqsel", "option10", language$, Option1(0).Caption, Option1(0), default_font)
    Option1(4).Caption = get_thing("liqsel", "option14", language$, Option1(4).Caption, Option1(4), default_font)
    set_fontstuff Text1, default_font
    Command1.Caption = gpps2("liqsel", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    Command2.Caption = gpps2("liqsel", "command2", language$, Command2.Caption)
    set_fontname Command2, default_font
    
    ' Other text
    ts$(1) = gpps2("liqsel", "ts1", language$, "Error: incorrect viscosity value")
    ts$(2) = gpps2("liqsel", "ts2", language$, "Error: incorrect liquid name value")
    ts$(3) = gpps2("liqsel", "ts3", language$, "Unit")
    ts$(4) = gpps2("liqsel", "ts4", language$, "Water")
    
End Sub
