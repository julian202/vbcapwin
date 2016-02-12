VERSION 5.00
Begin VB.Form Manfile 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manual Data File Creation"
   ClientHeight    =   6765
   ClientLeft      =   975
   ClientTop       =   525
   ClientWidth     =   9885
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
   ScaleHeight     =   6765
   ScaleWidth      =   9885
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.Frame Frame1 
         BackColor       =   &H0000FF00&
         Caption         =   "Test Type"
         Height          =   735
         Left            =   420
         TabIndex        =   23
         Top             =   240
         Width           =   9015
         Begin VB.OptionButton Option1 
            BackColor       =   &H0000FF00&
            Caption         =   "Capillary Flow"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Value           =   -1  'True
            Width           =   2895
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0000FF00&
            Caption         =   "Bubble Point"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   2
            Top             =   360
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0000FF00&
            Caption         =   "Permeability"
            Height          =   255
            Index           =   2
            Left            =   6240
            TabIndex        =   3
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   5
         Left            =   7680
         TabIndex        =   9
         Text            =   " "
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   0
         Left            =   2460
         TabIndex        =   4
         Top             =   1380
         Width           =   1515
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   6
         Left            =   7680
         TabIndex        =   10
         Text            =   " "
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   1
         Left            =   2460
         TabIndex        =   5
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   10
         Left            =   7680
         TabIndex        =   11
         Text            =   " "
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   2
         Left            =   2460
         TabIndex        =   6
         Top             =   2220
         Width           =   2175
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   7
         Left            =   7680
         TabIndex        =   12
         Text            =   " "
         Top             =   2340
         Width           =   975
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   3
         Left            =   2460
         TabIndex        =   7
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   8
         Left            =   7680
         TabIndex        =   13
         Text            =   "  "
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   4
         Left            =   2460
         TabIndex        =   8
         Text            =   " "
         Top             =   3060
         Width           =   1515
      End
      Begin VB.TextBox MFile 
         Height          =   315
         Index           =   9
         Left            =   7680
         TabIndex        =   14
         Text            =   " "
         Top             =   3180
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   5460
         TabIndex        =   22
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox Flow 
         Height          =   375
         Left            =   540
         TabIndex        =   15
         Top             =   4740
         Width           =   855
      End
      Begin VB.TextBox Pressure 
         Height          =   375
         Left            =   1620
         TabIndex        =   16
         Top             =   4740
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Default         =   -1  'True
         Height          =   375
         Left            =   3060
         TabIndex        =   17
         Top             =   4740
         Width           =   1335
      End
      Begin VB.CommandButton Helpcomm 
         Caption         =   "&Help"
         Height          =   375
         Left            =   1860
         TabIndex        =   18
         Top             =   5640
         Width           =   1455
      End
      Begin VB.CommandButton Saveas 
         Caption         =   "&Save As.."
         Height          =   375
         Left            =   3420
         TabIndex        =   19
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton Cancelcom 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear &Data"
         Height          =   375
         Left            =   6540
         TabIndex        =   21
         Top             =   5640
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Sample ID"
         Height          =   255
         Index           =   5
         Left            =   5100
         TabIndex        =   37
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Date of Test"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   1380
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Diameter (cm)"
         Height          =   255
         Index           =   6
         Left            =   5100
         TabIndex        =   35
         Top             =   1500
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "End User"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Cyl. Length or 0 (cm)"
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   33
         Top             =   1920
         Width           =   2595
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Test Reference"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   2220
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Thickness (cm)"
         Height          =   255
         Index           =   7
         Left            =   5100
         TabIndex        =   31
         Top             =   2340
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "S.Tension dynes/cm"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Zero Flow (cc/min)"
         Height          =   255
         Index           =   8
         Left            =   5100
         TabIndex        =   29
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Fluid Name"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   3060
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "Zero pressure"
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   27
         Top             =   3180
         Width           =   2595
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   " Enter Wet Flow, Pressure pairs going up then 0,0 and then dry flow and pressure going down."
         Height          =   915
         Left            =   300
         TabIndex        =   26
         Top             =   3720
         Width           =   4995
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "cc/min"
         Height          =   255
         Left            =   300
         TabIndex        =   25
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   1620
         TabIndex        =   24
         Top             =   5160
         Width           =   975
      End
   End
End
Attribute VB_Name = "Manfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
DefSng A-Z
Dim ZeroCnt As Integer
Dim ts$(16)                 ' Text strings for this form

Private Sub Cancelcom_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

    Dim lastp As Single
    
    If myVal(MFile(8).Text) < 0 Then
        MFile(8).SetFocus
        Exit Sub
    End If
    If myVal(MFile(9).Text) < 0 Then
        MFile(9).SetFocus
        Exit Sub
    End If
    If Flow.Text = "" Or Pressure.Text = "" Then
        Flow.SetFocus
        Exit Sub
    End If
    If List1.ListCount = 0 Then lastp = myVal(MFile(9).Text)
    If myVal(Flow.Text) = 0 And myVal(Pressure.Text) = 0 And (Option1(0).value Or Option1(2).value) Then
        List1.AddItem Flow.Text + "   <" + Chr$(124) + ">   " + Pressure.Text
        If Option1(2).value Then Command1.Enabled = False
        If Option1(0).value Then ZeroCnt = ZeroCnt + 1: lastp = 9999
        If ZeroCnt >= 2 Then Command1.Enabled = False
        Pressure.Text = ""
        Flow.Text = ""
        Flow.SetFocus
        Exit Sub
    End If

    If ZeroCnt >= 1 Then
        If myVal(Pressure.Text) >= lastp - 0.00001 Then
            MsgBox ts$(1), 16, ts$(2)             '"Pressure must be lower than the last entered pressure."/"Data Error"
            Pressure.SetFocus
            Exit Sub
        End If
    ElseIf myVal(Pressure.Text) <= lastp + 0.00001 Then
        MsgBox ts$(3), 16, ts$(2)            ' "Pressure must be higher than the Zero or last entered pressure."/"Data Error"
        Pressure.SetFocus
        Exit Sub
    End If

    List1.AddItem Flow.Text + "   <" + Chr$(124) + ">   " + Pressure.Text
    lastp = myVal(Pressure.Text)
    Pressure.Text = ""
    Flow.Text = ""
    Flow.SetFocus
    If (List1.ListCount = 1) And (Option1(1).value) Then Command1.Enabled = False

End Sub

Private Sub Command2_Click()
    While List1.ListCount
        List1.RemoveItem 0
    Wend
End Sub

Private Sub Flow_GotFocus()
    Flow.SelStart = 0
    Flow.SelLength = Len(Flow.Text)
End Sub

Private Sub Flow_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, Asc(Decimal_Point$), 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Load()

    'Dim i As Integer
    
    LoadTextStrings
    TitleScrn.modifymenu(1).Enabled = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    'RemoveSysMenu Me
    MFile(0).Text = date$
    Label4.Caption = PU$
    Label1(9).Caption = ts$(4) + " (" + PU$ + ")"               ' "Zero pressure"
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(5) + Str$(unitnumber)   '"Unit"
    End If
    
    ' Superseded by LoadTextStrings:
    'For i = 0 To 10
    '    MFile(i).FontName = system_font
    '    MFile(i).fontsize = font_size
    '    MFile(i).fontbold = font_bold
    'Next i
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    TitleScrn.modifymenu(1).Enabled = True
End Sub



Private Sub Helpcomm_Click()
    Dim r As Long
    r = WinHelp(hwnd, HelpFile$, Help_Context, ByVal 25&)
End Sub



Private Sub MFile_Change(Index As Integer)
    If (Index = 8 Or Index = 9) And List1.ListCount > 0 Then
        MsgBox ts$(6), 16, ts$(7)        ' "Changing this value will invalidate any data."/"Information"
    End If
End Sub

Private Sub MFile_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If (Index = 3) Or (Index = 6) Or (Index = 7) Or (Index = 8) Or (Index = 9) Or (Index = 10) Then
        Select Case KeyAscii
            Case 8, Asc(Decimal_Point$), 48 To 57
            Case Else
                KeyAscii = 0
        End Select
    End If

End Sub

Private Sub MFile_LostFocus(Index As Integer)
    
    If (Index = 3) Or (Index = 6) Or (Index = 7) Or (Index = 8) Or (Index = 9) Or (Index = 10) Then
        MFile(Index).Text = Format$(myVal(MFile(Index).Text), "####0.0000")
    End If
    If (Index = 8) And (myVal(MFile(8).Text) < 0) Then MFile(8).SetFocus

End Sub

Private Sub Option1_Click(Index As Integer)
    
    reply% = MsgBox(ts$(8), 1, ts$(9))     ' "Changing the Test Type will clear any entered data."/"Test Type"
    If reply% = 2 Then Exit Sub

    List1.clear
    ZeroCnt = 0
    Command1.Enabled = True
    MFile(8).Enabled = True
    MFile(9).Enabled = True
    MFile(8).Text = ""
    MFile(9).Text = ""

    If Index = 1 Then
        MFile(3).Visible = True
        Label1(3).Visible = True
        MFile(6).Visible = False
        Label1(6).Visible = False
        MFile(10).Visible = False
        Label1(10).Visible = False
        MFile(7).Visible = False
        Label1(7).Visible = False
        Label2.Caption = ts$(10)        ' "Enter Bubble Point Flow, Pressure."
    ElseIf Index = 2 Then
        MFile(3).Visible = False
        Label1(3).Visible = False
        MFile(6).Visible = True
        MFile(7).Visible = True
        Label1(6).Visible = True
        Label1(7).Visible = True
        MFile(10).Visible = True
        Label1(10).Visible = True
        Label2.Caption = ts$(11)        ' "Enter Flow, Pressure pairs going up, then '0,0' to end."
    Else
        If Option1(0).value Then Label2.Caption = ts$(12)       ' "Enter Wet Flow, Pressure pairs going up then 0,0 and then dry flow and pressure going down, then 0,0 to end."
        MFile(3).Visible = True
        MFile(4).Visible = True
        Label1(3).Visible = True
        Label1(4).Visible = True
        MFile(6).Visible = True
        MFile(7).Visible = True
        Label1(6).Visible = True
        Label1(7).Visible = True
        MFile(10).Visible = True
        Label1(10).Visible = True
    End If

End Sub

Private Sub Pressure_GotFocus()
    Pressure.SelStart = 0
    Pressure.SelLength = Len(Pressure.Text)
End Sub

Private Sub Pressure_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 8, Asc(Decimal_Point$), 48 To 57
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub Saveas_Click()

    Dim i%, temp$
    Dim P As Single
    Dim qq As Single
    Dim qq2 As Single
    Dim Q0 As Single
    Dim Q As Single
    Dim barpos As Integer
    Dim fn As Integer
    
    For i% = 0 To 9
        If (i% <> 1) And (i% <> 2) Then
            If (Len(MFile(i%).Text) <= 1) And (MFile(i%).Visible) Then
                MsgBox ts$(13) + " " + Label1(i%).Caption + " .", 0, ts$(14)   ' "Please make entries for"/"Enter Data"
                MFile(i%).SetFocus
                Exit Sub
            End If
        End If
    Next i%
    If (Option1(1).value) And (List1.ListCount <> 1) Then
        MsgBox ts$(15), 0, ts$(2)          ' "Please enter the bubble flow and pressure."/"Data Error"
        Flow.SetFocus
        Exit Sub
    End If

    fsel_name$ = ""
    fsel_title$ = ts$(16)                   ' "New Data File."
    fsel_path$ = path(0)
    fsel_io = False
    fsel Me.hwnd
    If fsel_return$ = "" Then
        Exit Sub
    End If
    path(0) = fsel_path$
    file$ = fsel_return$

    fn = FreeFile
    Open file$ For Output As #fn
        Print #fn, MFile(0).Text
        Print #fn, MFile(1).Text
        Print #fn, MFile(2).Text
        If Option1(2).value Then
            MFile(3).Text = "0"
        End If
        Print #fn, MFile(4).Text
        If Option1(0).value Then
            Print #fn, Str$(-myVal(MFile(3).Text))   'SURFTEN(current_unit%)
            qq = myVal(MFile(6).Text)
            qq2 = myVal(MFile(10).Text)
            If qq2 = 0 Then
                Print #fn, Str$(qq)
            Else
                Print #fn, 0
                Print #fn, Str$(qq); ","; Str$(qq2)
            End If
            Print #fn, Str$(myVal(MFile(7).Text))
        Else
            Print #fn, Str$(myVal(MFile(3).Text))   'SURFTEN(current_unit%)
        End If

        Print #fn, MFile(5).Text

        If Option1(2).value Then
            qq = myVal(MFile(6).Text)
            qq2 = myVal(MFile(10).Text)
            If qq2 = 0 Then
                Print #fn, Str$(qq)
            Else
                Print #fn, 0
                Print #fn, Str$(qq); ","; Str$(qq2)
            End If
            Print #fn, Str$(myVal(MFile(7).Text))
            Print #fn, Str$(myVal(MFile(8).Text)); ","; Str$(myVal(MFile(9).Text) / PCNV)
            Q0 = myVal(MFile(8).Text)
            P = myVal(MFile(9).Text)
            For i% = 0 To List1.ListCount - 1
                temp$ = LTrim$(RTrim$(List1.List(i%)))
                barpos = InStr(temp$, Chr$(124))
                Q = myVal(Mid$(temp$, 1, barpos - 2))
                P = myVal(Mid$(temp$, barpos + 2)) / PCNV
                Print #fn, Str$(Q - Q0); ","; Str$(P)
            Next i%
        Else
            Print #fn, Str$(myVal(MFile(8).Text)); ","; Str$(myVal(MFile(9).Text) / PCNV)
            Q = myVal(MFile(8).Text)
            P = myVal(MFile(9).Text)
            For i% = 0 To List1.ListCount - 1
                temp$ = LTrim$(RTrim$(List1.List(i%)))
                barpos = InStr(temp$, Chr$(124))
                Q = myVal(Mid$(temp$, 1, barpos - 2))
                P = myVal(Mid$(temp$, barpos + 2))
                Print #fn, Str$(Q - Q0); ","; Str$(P / PCNV)
            Next i%
            Print #fn, "0,0"
        End If
        Close #fn
        Flow.Text = ""
        Pressure.Text = ""
        List1.clear
        Command1.Enabled = True
        Me.SetFocus
        ZeroCnt = 0
        MFile(8).Enabled = True
        MFile(9).Enabled = True

End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    Manfile.Caption = gpps2("manfile", "window title", language$, Manfile.Caption)
    Frame1.Caption = get_thing("manfile", "frame1", language$, Frame1.Caption, Frame1, default_font)
    For i = 0 To 2
        Option1(i).Caption = get_thing("manfile", "option1" + Str$(i), language$, Option1(i).Caption, Option1(i), default_font)
    Next i
    For i = 0 To 10
        Label1(i) = get_thing("manfile", "label1" + Str$(i), language$, Label1(i).Caption, Label1(i), default_font)
        set_fontstuff MFile(i), default_font
    Next i
    Label2.Caption = get_thing("manfile", "label2", language$, Label2.Caption, Label2, default_font)
    Label3.Caption = get_thing("manfile", "label3", language$, Label3.Caption, Label3, default_font)
    Command1.Caption = gpps2("manfile", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    Command2.Caption = gpps2("manfile", "command2", language$, Command2.Caption)
    set_fontname Command2, default_font
    Helpcomm.Caption = gpps2("manfile", "help", language$, Helpcomm.Caption)
    set_fontname Helpcomm, default_font
    Saveas.Caption = gpps2("manfile", "saveas", language$, Saveas.Caption)
    set_fontname Saveas, default_font
    Cancelcom.Caption = gpps2("manfile", "cancel", language$, Cancelcom.Caption)
    set_fontname Cancelcom, default_font
    set_fontstuff List1, default_font
    set_fontstuff Flow, default_font
    set_fontstuff Pressure, default_font
    set_fontstuff Label4, default_font
    
    ' Other text
    ts$(1) = gpps2("manfile", "ts1", language$, "Pressure must be lower than the last entered pressure.")
    ts$(2) = gpps2("manfile", "ts2", language$, "Data Error")
    ts$(3) = gpps2("manfile", "ts3", language$, "Pressure must be higher than the Zero or last entered pressure.")
    ts$(4) = gpps2("manfile", "ts4", language$, "Zero pressure")
    ts$(5) = gpps2("manfile", "ts5", language$, "Unit")
    ts$(6) = gpps2("manfile", "ts6", language$, "Changing this value will invalidate any data.")
    ts$(7) = gpps2("manfile", "ts7", language$, "Information")
    ts$(8) = gpps2("manfile", "ts8", language$, "Changing the test type will clear any entered data.")
    ts$(9) = gpps2("manfile", "ts9", language$, "Test Type")
    ts$(10) = gpps2("manfile", "ts10", language$, "Enter Bubble Point Flow, Pressure.")
    ts$(11) = gpps2("manfile", "ts11", language$, "Enter Flow, Pressure pairs going up, then '0,0' to end.")
    ts$(12) = gpps2("manfile", "ts12", language$, "Enter Wet Flow, Pressure pairs going up, then '0,0', and then dry flow and pressure going down, then '0,0' to end.")
    ts$(13) = gpps2("manfile", "ts13", language$, "Please make entries for")
    ts$(14) = gpps2("manfile", "ts14", language$, "Enter Data")
    ts$(15) = gpps2("manfile", "ts15", language$, "Please enter the bubble flow and pressure.")
    ts$(16) = gpps2("manfile", "ts16", language$, "New Data File")
    
End Sub

