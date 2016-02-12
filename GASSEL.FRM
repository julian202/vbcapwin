VERSION 5.00
Begin VB.Form GasSel 
   Caption         =   "Gas Select"
   ClientHeight    =   3330
   ClientLeft      =   1095
   ClientTop       =   1485
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
   ScaleHeight     =   3330
   ScaleWidth      =   4905
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2280
      TabIndex        =   9
      Text            =   " "
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2280
      TabIndex        =   7
      Text            =   " "
      Top             =   2280
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "X: Other gas"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   3975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "R: Argon"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "N: Nitrogen"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "H: Helium"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "A: Air"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Viscosity (cP):"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Name of Gas:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Select Gas:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "GasSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ts$(7)                  ' Text strings for this form

Private Sub Command1_Click()

Dim a$, b$
Dim x As Single

If Option1(0).Value Then
    GasID$ = "A"
    Gas$ = ts$(4)       ' "Air"
    Unload Me
ElseIf Option1(1).Value Then
    GasID$ = "H"
    Gas$ = ts$(5)       ' "Helium"
    Unload Me
ElseIf Option1(2).Value Then
    GasID$ = "N"
    Gas$ = ts$(6)       ' "Nitrogen"
    Unload Me
ElseIf Option1(3).Value Then
    GasID$ = "R"
    Gas$ = ts$(7)      ' "Argon"
    Unload Me
Else
    If myVal(Text2.Text) <= 0 Then
        MsgBox ts$(1)           ' "Error: incorrect viscosity value"
        Exit Sub
    End If
    If Len(Text1.Text) = 0 Then
        MsgBox ts$(2)           ' "Error: incorrect gas name value"
        Exit Sub
    End If
    a$ = Text1.Text
    x = myVal(Text2.Text)
    b$ = LTrim$(Str$(x))
    Gas$ = Left$(b$ + Space$(20), 20) + a$
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
        Me.Caption = Me.Caption + " - " + ts$(3) + Str$(unitnumber)     ' "Unit"
    End If
    ' Superseded by LoadTextStrings:
    'Text1.FontName = system_font
    'Text1.fontsize = font_size
    'Text1.fontbold = font_bold
    'Text2.FontName = system_font
    'Text2.fontsize = font_size
    'Text2.fontbold = font_bold
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
    
    Dim i As Integer
    
    ' Form elements
    GasSel.Caption = gpps2("gassel", "window title", language$, GasSel.Caption)
    Label1.Caption = get_thing("gassel", "label1", language$, Label1.Caption, Label1, default_font)
    Label2.Caption = get_thing("gassel", "label2", language$, Label2.Caption, Label2, default_font)
    Label3.Caption = get_thing("gassel", "label3", language$, Label3.Caption, Label3, default_font)
    For i = 0 To 4
        Option1(i).Caption = get_thing("gassel", "Option1" + Str$(i), language$, Option1(i).Caption, Option1(i), default_font)
    Next i
    set_fontstuff Text1, default_font
    Command1.Caption = gpps2("gassel", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    Command2.Caption = gpps2("gassel", "command2", language$, Command2.Caption)
    set_fontname Command2, default_font
    
    ' Other text
    ts$(1) = gpps2("gassel", "ts1", language$, "Error: incorrect viscosity value")
    ts$(2) = gpps2("gassel", "ts2", language$, "Error: incorrect gas name value")
    ts$(3) = gpps2("gassel", "ts3", language$, "Unit")
    ts$(4) = gpps2("gassel", "ts4", language$, "Air")
    ts$(5) = gpps2("gassel", "ts5", language$, "Helium")
    ts$(6) = gpps2("gassel", "ts6", language$, "Nitrogen")
    ts$(7) = gpps2("gassel", "ts7", language$, "Argon")
    
End Sub
