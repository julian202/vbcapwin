VERSION 5.00
Begin VB.Form progress 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Current Test Status"
   ClientHeight    =   6750
   ClientLeft      =   1950
   ClientTop       =   1770
   ClientWidth     =   9495
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
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6750
   ScaleWidth      =   9495
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   9255
      Begin VB.PictureBox PictureBP 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         DrawWidth       =   2
         Height          =   4335
         Left            =   1200
         ScaleHeight     =   4275
         ScaleWidth      =   7755
         TabIndex        =   21
         Top             =   690
         Visible         =   0   'False
         Width           =   7815
      End
      Begin VB.PictureBox sadface 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   360
         Picture         =   "PROGRESS.frx":0000
         ScaleHeight     =   480
         ScaleMode       =   0  'User
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   1920
         Width           =   500
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         DrawWidth       =   2
         Height          =   4335
         Left            =   1200
         ScaleHeight     =   4275
         ScaleWidth      =   7755
         TabIndex        =   4
         Top             =   690
         Width           =   7815
      End
      Begin VB.Timer Timer1 
         Left            =   600
         Top             =   5760
      End
      Begin VB.ListBox List1 
         Height          =   3960
         Left            =   1200
         TabIndex        =   3
         Top             =   690
         Width           =   7815
      End
      Begin VB.PictureBox happyface 
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   360
         Picture         =   "PROGRESS.frx":0C5E
         ScaleHeight     =   495
         ScaleMode       =   0  'User
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   1440
         Width           =   500
      End
      Begin VB.Label lblMicroflowVolume 
         BackColor       =   &H00FFFF80&
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   18
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Y_Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "YMax"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "cc/min"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Y_Zero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Yzero"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label X_Zero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Xzero"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "PU$"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4650
         TabIndex        =   13
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label X_Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Xmax"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8160
         TabIndex        =   12
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label W_Curve 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1350
         TabIndex        =   11
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Label Line25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   5880
         Width           =   9255
      End
      Begin VB.Label Line26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   6120
         Width           =   9255
      End
      Begin VB.Label FileNmLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4680
         TabIndex        =   8
         Top             =   165
         Width           =   75
      End
      Begin VB.Label SampleIDLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4710
         TabIndex        =   7
         Top             =   450
         Width           =   75
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   6
         Top             =   5520
         Width           =   3495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Label4"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   5160
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Pause"
      Index           =   0
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Next S&tep"
      Enabled         =   0   'False
      Index           =   1
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Abort"
      Enabled         =   0   'False
      Index           =   2
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Manual"
      Enabled         =   0   'False
      Index           =   3
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Record"
      Index           =   4
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Bubble"
      Index           =   5
   End
   Begin VB.Menu mainmenu 
      Caption         =   "&Help"
      Index           =   6
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Stability Debug"
      Index           =   7
   End
   Begin VB.Menu mainmenu 
      Caption         =   "OTHER CHAMBER READY"
      Index           =   8
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim ts$(6)              ' Text strings for this form

Private Sub Form_KeyPress(KeyAscii As Integer)
    user_keypress = KeyAscii
End Sub

Private Sub Form_Load()

    LoadTextStrings
    'WindowState = 2
    Me.X_Zero.Caption = "0.0"
    Me.Y_Zero.Caption = "0.0"
    Me.Label2.Caption = PU$
    Me.Y_Max.Caption = LTrim$(RTrim$(Format$(STARTF)))
    Me.X_Max.Caption = LTrim$(RTrim$(Format$(STARTP * PCNV, "#######0.00")))
    Me.Picture1.Scale (0, STARTF)-(STARTP, 0)
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.top = (Screen.Height - Me.Height) / 2
    If TType%(current_unit%) >= 4 Then
        mainmenu(1).Visible = False
        mainmenu(4).Visible = False
        mainmenu(5).Visible = False
    End If
    mainmenu(1).Enabled = False
    mainmenu(2).Enabled = False
    mainmenu(3).Enabled = False
    If Not supervisor Then
        mainmenu(3).Visible = False
    End If
    SampleIDLabel.Caption = ts$(1) + ": " & sid$(current_unit%)             ' "Sample ID"
    FileNmLabel.Caption = ts$(2) + ":  " & OutFilename$(current_unit%)      '"Filename"
    Me.Timer1.Enabled = False
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(3) + str$(unitnumber)    '"Unit"
    End If
    
    Label4.Visible = False
    stability_debug = False
    mainmenu(7).Visible = TitleScrn.debugmenu(1).Checked    ' only show debugging menu if enabled
                                                            ' from main
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
    'added 11-29-07 --Denis
    If gpps2("Capstuff", "sequential_testing_enabled", CSFile$, "N") = "Y" Then
        mainmenu(8).Visible = True
    Else
        mainmenu(8).Visible = False
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    Dim i As Integer
    Dim Ret$
    Ret$ = String$(255, " ")

    Me.Timer1.Enabled = False
End Sub

Private Sub mainmenu_Click(Index As Integer)

    Dim OLFlow%, r As Long
    Dim OPA As Integer, OPB As Integer
    Dim OPres%, OHFlow%
    Dim Answer%
    ' 6.71.20 new variable
    Dim OSv10 As Boolean
    
    Select Case Index
        Case 0
            ' if we are holding, we don't want to
            ' if we are not holding, we want to so...
            want_to_hold = Not holding
            ' the rest of this has been moved.  Hold
            ' doesn't happen until it is safe for it
            ' to do so.
        Case 1
            HKey2$ = "T"
            want_to_hold = False
        Case 2
            Answer = MsgBox(ts$(4), vbYesNo, ts$(5))
            If Answer = vbYes Then     ' "Do you really want to abort this test?"/"Abort Test"
              HKey$ = "A"
              Aborted = True
              
              want_to_hold = False
            End If
        Case 3
          If holding Then
            mainmenu(0).Enabled = False
            mainmenu(1).Enabled = False
            mainmenu(2).Enabled = False
            mainmenu(3).Enabled = False
            mainmenu(4).Enabled = False
            mainmenu(5).Enabled = False
            mainmenu(6).Enabled = False

            OPres% = Pres%
            OHFlow% = HFLOW%
            OLFlow% = lflow%
            RUNNING = True
            intest = False
            ' 6.71.20
            OSv10 = suspend_v10
            If hasMultipleMVs Then
                ManualControl1.Show 0
            Else
                ManualControl.Show 0
            End If
            RUNNING = True
            run_manual_control
            intest = True
            ' 6.71.20
            suspend_v10 = OSv10
            ' 6.71.75
            Pres% = OPres%
            HFLOW% = OHFlow%
            lflow% = OLFlow%

            mainmenu(0).Enabled = True
            mainmenu(1).Enabled = True
            mainmenu(2).Enabled = True
            mainmenu(3).Enabled = True
            mainmenu(4).Enabled = True
            mainmenu(5).Enabled = True
            mainmenu(6).Enabled = True
          End If
        Case 4
            HKey$ = "R"
        Case 5
            HKey$ = "B"
        Case 6
            r = WinHelp(hwnd, HelpFile$, Help_Index, ByVal 50&)
        Case 7
            stabilitydebug.Show 0
        'added 11-29-07 --Denis
        Case 8
            'something about the sequential test
            mainmenu(8).Visible = False
            
            'alternate the chambers that are ready
            If chamber_ready = 1 Then
                chamber_ready = 2
            ElseIf chamber_ready = 2 Then
                chamber_ready = 1
            ElseIf chamber_ready = 0 Then
                If tempChamberVal = 1 Then
                    chamber_ready = 2
                ElseIf tempChamberVal = 2 Then
                    chamber_ready = 1
                End If
            End If
    End Select
    
End Sub

Private Sub Timer1_Timer()

Dim u$, lt As Single

If multiChamberSystem = False Or manualMultiChamber = True Then
    u$ = ""
Else
    u$ = ts$(3) + str$(current_unit%)       ' "Unit"
End If
    lt = Timer
    If lt + 0.1 < RunTimer Then
        RunTimer = RunTimer - 86400
    End If
    Me.Caption = ts$(6) + " " + u$ + "  " + Format$((lt - RunTimer) / 86400, "hh:mm:ss") ' "Current Test Status"
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(3) + str$(unitnumber)    ' "Unit"
    End If
    
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    progress.Caption = gpps2("progress", "window title", language$, progress.Caption)
    For i = 0 To 6
        mainmenu(i).Caption = gpps2("progress", "mainmenu" + str$(i), language$, mainmenu(i).Caption)
    Next i
    set_fontstuff FileNmLabel, default_font
    set_fontstuff SampleIDLabel, default_font
    Label1.Caption = get_thing("progress", "label1", language$, Label1.Caption, Label1, default_font)
    set_fontstuff Label2, default_font
    For i = 0 To 2
        set_fontstuff Label3(i), default_font
    Next i
    set_fontstuff Label4, default_font
    set_fontstuff Line25, default_font
    set_fontstuff Line26, default_font
    set_fontstuff List1, default_font
    set_fontstuff W_Curve, default_font
    set_fontstuff X_Max, default_font
    set_fontstuff Y_Max, default_font
    set_fontstuff X_Zero, default_font
    set_fontstuff Y_Zero, default_font
    
    ' Other text
    ts$(1) = gpps2("progress", "ts1", language$, "Sample ID")
    ts$(2) = gpps2("progress", "ts2", language$, "Filename")
    ts$(3) = gpps2("progress", "ts3", language$, "Unit")
    ts$(4) = gpps2("progress", "ts4", language$, "Do you really want to abort this test?")
    ts$(5) = gpps2("progress", "ts5", language$, "Abort Test")
    ts$(6) = gpps2("progress", "ts6", language$, "Current Test Status")
    
End Sub

