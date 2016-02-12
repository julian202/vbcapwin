VERSION 5.00
Begin VB.Form qcmain 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PMI Porometer QC Mode"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   2055
   ClientWidth     =   9525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Selected Test"
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run Selected Test"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9255
   End
   Begin VB.Menu qcmenu 
      Caption         =   "Mode"
      Begin VB.Menu exitqcmodemenu 
         Caption         =   "Exit QC Mode"
      End
      Begin VB.Menu exitprogrammenu 
         Caption         =   "Exit Program"
      End
   End
   Begin VB.Menu gmenu 
      Caption         =   "Test"
      Begin VB.Menu selectmenu 
         Caption         =   "Select Test"
      End
   End
End
Attribute VB_Name = "qcmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim startindex As Integer
Dim numgroups As Integer
Dim ts$(3)             ' Text strings for this form

Private Sub Command2_Click()

' run auto test
Init_For_Ctrl True
current_unit% = 1
If multiChamberSystem And (manualMultiChamber = False) Then
    Do Until selchamber(current_unit%)
        current_unit% = current_unit% + 1
        If current_unit% > chambers Then
            MsgBox ts$(2)                       ' ("Error:  This test does not have any enabled chambers")
            Exit Sub
        End If
    Loop
End If
' changed to True - surik
 RUNNING = True
' RUNNING = False ' this will be set to true if we are really
                ' running a test
If ExtraPG Then Pres% = 2 Else Pres% = 1
intest = True
save_setup_data_flag = False
first_test_setup = True
Me.Hide
pleasewait.Show 0
Testscrn.Show 1
' if save_setup_data_flag is true but RUNNING is false,
' these routines will exit after saving the information
If save_setup_data_flag Then
    If multiChamberSystem And (manualMultiChamber = False) Then
        RunMultipleTest
    Else
        RunSingleTest
    End If
End If
Me.Show

End Sub

Private Sub exitprogrammenu_Click()
    Unload TitleScrn
    Unload Me
End Sub

Private Sub exitqcmodemenu_Click()

If superpass$ <> "" Then
    GetValue.Label1.Caption = ts$(3) + " :"     ' "Enter Password"
    GetValue.Text1.Text = ""
    GetValue.Text1.SelStart = 0
    GetValue.Text1.SelLength = 0
    GetValue.Label1.Tag = "text"
    GetValue.Continue.default = True
    GetValue.Show 1
    GetValue.Label1.Tag = ""
    If Got_Value = -9 Then Exit Sub
    If Got_Text <> superpass$ Then Exit Sub
End If
simpleqc_enable = False
WPPS "Capstuff", "SimpleQC", "0", CSFile$
TitleScrn.Show

TitleScrn.WindowState = 0
Unload Me

End Sub

Private Sub Form_Load()

    LoadTextStrings
    Label1.Caption = Curr_U$
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor

End Sub

Private Sub selectmenu_Click()

    T_Select$ = "USER"
    Selection.Show 1
    Label1.Caption = Curr_U$
    WPPS "default", "user", Curr_U$, IFile$
    load_user_global_stuff
    
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    qcmain.Caption = gpps2("qcmain", "window title", language$, qcmain.Caption)
    qcmenu.Caption = gpps2("qcmain", "menu", language$, qcmenu.Caption)
    exitprogrammenu.Caption = gpps2("qcmain", "exit", language$, exitprogrammenu.Caption)
    exitqcmodemenu.Caption = gpps2("qcmain", "exitqcmode", language$, exitqcmodemenu.Caption)
    gmenu.Caption = gpps2("qcmain", "groupmenu", language$, gmenu.Caption)
    selectmenu.Caption = gpps2("qcmain", "selectmenu", language$, selectmenu.Caption)
    Command2.Caption = get_thing("qcmain", "command2", language$, Command2.Caption, Command2, default_font)
    Label1.Caption = get_thing("qcmain", "label1", language$, Label1.Caption, Label1, default_font)
    Label1.fontsize = Label1.fontsize + 10
    Label1.ForeColor = vbBlue
    Frame1.Caption = get_thing("qcmain", "label2", language$, Frame1.Caption, Frame1, default_font)
    Frame1.fontsize = Frame1.fontsize + 4

    ' Other text
 '   ts$(1) = gpps2("qcmain", "ts1", language$, "Unused")
    ts$(2) = gpps2("qcmain", "ts2", language$, "Error:  This test does not have any enabled chambers")
    ts$(3) = gpps2("qcmain", "ts3", language$, "Enter Password")
    
End Sub
