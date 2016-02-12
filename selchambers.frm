VERSION 5.00
Begin VB.Form selchambers 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Chambers"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox sequential_testing 
         Caption         =   "Sequential chamber testing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 3"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 4"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 5"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 6"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 7"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 8"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 9"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CheckBox chambercheck 
         Caption         =   "Chamber 10"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   3480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "selchambers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim localSelectedChamber As Integer

Private Sub chambercheck_Click(Index As Integer)
Dim i As Integer
Dim temp_dry_chambers As Integer
Dim temp_manualMultiChamber As Boolean
    If manualMultiChamber Or (dry_chambers > 1) Then
        temp_dry_chambers = dry_chambers
        temp_manualMultiChamber = manualMultiChamber
        localSelectedChamber = Index
        manualMultiChamber = False ' temporary so next lines don't cause recursive loop
        dry_chambers = 1
        For i = 1 To 10
            chambercheck(i) = 0
        Next i
        chambercheck(localSelectedChamber) = 1
        ' restore
        manualMultiChamber = temp_manualMultiChamber
        dry_chambers = temp_dry_chambers
    End If
    
    'added 11-29-07 --Denis
    'Make sure that when the sequential_testing is selected that the 2 chambers cant be played
    'around with to cause errors. Make the chambers grayed.
    If sequential_testing.value = 1 Then
        chambercheck(1).value = 2
        chambercheck(2).value = 2
    End If

End Sub

Private Sub Command1_Click()

Dim i As Integer

If manualMultiChamber Or (dry_chambers > 1) Then
    manuallySelectedChamber = localSelectedChamber
    WPPS Curr_U$, "manually_selected_chamber", Str$(manuallySelectedChamber), IFile$
Else
    For i = 1 To chambers
        selchamber(i) = (chambercheck(i).value >= 1)
        WPPS Curr_U$, "chamber select" + Str$(i), Str$(chambercheck(i)), IFile$
    Next i
End If
Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

Dim i As Integer
Dim temp_dry_chambers As Integer
Dim temp_manualMultiChamber As Boolean

LoadTextStrings
If manualMultiChamber Or (dry_chambers > 1) Then
    temp_dry_chambers = dry_chambers
    temp_manualMultiChamber = manualMultiChamber
    For i = 1 To chambers
        chambercheck(i).value = 0
    Next i
    chambercheck(manuallySelectedChamber) = 1
    localSelectedChamber = manuallySelectedChamber
    ' restore
    manualMultiChamber = temp_manualMultiChamber
    dry_chambers = temp_dry_chambers
Else
    For i = 1 To chambers
        chambercheck(i).value = IIf(selchamber(i), 1, 0)
    Next i
End If
If chambers < 10 Then
    For i = chambers + 1 To 10
        chambercheck(i).Visible = False
    Next i
End If
'edc 12-11-06 alter border color and caption
    'Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
      
      
    'added 11-29-07 --Denis
    If gpps2("Capstuff", "sequential_testing_enabled", CSFile$, "N") = "Y" Then
        sequential_testing.value = val(gpps2(Curr_U$, "sequential_testing", IFile$, "0"))
        If sequential_testing.value = 1 Then
            'set chamber 1 and 2 to grayed(selected value)
            chambercheck(1).value = 2
            chambercheck(2).value = 2
        End If
    Else
        sequential_testing.Visible = False
    End If
    
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    selchambers.Caption = gpps2("selchambers", "window title", language$, selchambers.Caption)
    For i = 1 To 10
        chambercheck(i).Caption = get_thing("selchambers", "chambercheck" + Format$(i), language$, chambercheck(i).Caption, chambercheck(i), default_font)
    Next i
    Command1.Caption = gpps2("selchambers", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    Command2.Caption = gpps2("selchambers", "command2", language$, Command2.Caption)
    set_fontname Command2, default_font

    
End Sub

'Added 11-29-07 --Denis
Private Sub sequential_testing_Click()
    If sequential_testing.value = 0 Then
        'restore to selected chambers
        chambercheck(1).value = 1
        chambercheck(2).value = 1
        sequentialTesting = False
    End If
    If sequential_testing.value = 1 Then
        'set chamber 1 and 2 to grayed(selected value)
        chambercheck(1).value = 2
        chambercheck(2).value = 2
        sequentialTesting = True
    End If
    
    WPPS Curr_U$, "sequential_testing", Str$(sequential_testing.value), IFile$
End Sub
