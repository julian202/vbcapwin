VERSION 5.00
Begin VB.Form CalibBoard 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PC Interface Board Calibration"
   ClientHeight    =   3750
   ClientLeft      =   2250
   ClientTop       =   1920
   ClientWidth     =   6150
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
   ScaleHeight     =   3750
   ScaleWidth      =   6150
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5895
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   5160
         Top             =   1800
      End
      Begin VB.CommandButton bdexit 
         Appearance      =   0  'Flat
         Caption         =   "E&xit"
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Span"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zero"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Both"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   5
         Top             =   1440
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   10
         Left            =   240
         Max             =   255
         TabIndex        =   7
         Top             =   2400
         Value           =   10
         Width           =   5295
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         LargeChange     =   10
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   8
         Top             =   3000
         Value           =   10
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Span () = "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Zero () = "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Also make sure calibration voltage is set to 2 volts"
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   720
         TabIndex        =   2
         Top             =   960
         Width           =   4305
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delay: 10"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Average: 10"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   2535
      End
   End
End
Attribute VB_Name = "CalibBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim last_xignore As Byte
Dim last_xmult As Byte
Dim first_xignore As Byte
Dim first_xmult As Byte
Dim ts$(6)              ' Text strings for this form

Private Sub bdexit_Click()
    Rem don't unload when clicked as there may be a timer
    Rem function still going on
    Rem Unload me
    Rem simply disable this button.  The timer function will
    Rem check for this and then unload the form
    bdexit.Enabled = False
End Sub

Private Sub Form_Load()

    LoadTextStrings
    'RemoveSysMenu Me
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(1) + Str$(unitnumber)    ' "Unit"
    End If
    If Version < 7 Then
        Me.Height = 2885
        Frame1.Height = 2300
        Label4.Visible = False
        ' other things are below the edge of the form so they don't need to be made invisible
    Else
        Label4.Caption = ts$(2) + ":" + Str$(xignore)                       ' "Delay"
        HScroll1.value = xignore
        last_xignore = xignore
        first_xignore = xignore
        Label5.Caption = ts$(3) + ":" + Str$(xmult)                         ' "Average"
        HScroll2.value = xmult
        last_xmult = xmult
        first_xmult = xmult
    End If
    'edc 12-11-06 alter border color and caption
    'Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
If bdexit.Enabled Then
    bdexit.Enabled = False
    cancel = -1
End If
End Sub

Private Sub HScroll1_Change()
    xignore = HScroll1.value
    Label4.Caption = ts$(2) + ":" + Str$(xignore)       ' "Delay"
End Sub

Private Sub HScroll1_Scroll()
    xignore = HScroll1.value
    Label4.Caption = ts$(2) + ":" + Str$(xignore)       ' "Delay"
End Sub

Private Sub HScroll2_Change()
    xmult = HScroll2.value
    Label5.Caption = ts$(3) + ":" + Str$(xmult)         ' "Average"
End Sub

Private Sub HScroll2_Scroll()
    xmult = HScroll2.value
    Label5.Caption = ts$(3) + ":" + Str$(xmult)         ' "Average"
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False
Do
    If last_xignore <> xignore Then
        last_xignore = xignore
        Send_RS232b "i", xignore
    End If
    If last_xmult <> xmult Then
        last_xmult = xmult
        Send_RS232b "m", xmult
    End If
    If Option1(0).value Or Option1(2).value Then
        ReadXReturnX4 14
        Label1.Caption = ts$(4) + " (" + ts$(6) + Str$(DAC_two) + ") = " + Format$(x4, "####0")    '"Span"/"should be"
        readings_counter = readings_counter + 1
    End If
    If Option1(1).value Or Option1(2).value Then
        ReadXReturnX4 15
        Label2.Caption = ts$(5) + " (" + ts$(6) + Str$(DAC_zero) + ") = " + Format$(x4, "####0")   '"Zero"/"should be"
        readings_counter = readings_counter + 1
    End If
Loop Until bdexit.Enabled = False
If xignore <> first_xignore Then
    WPPS "Capstuff", "xignore", Str$(xignore), CSFile$
End If
If xmult <> first_xmult Then
    WPPS "Capstuff", "xmult", Str$(xmult), CSFile$
End If
'bdexit.Enabled = True
Unload Me

End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    CalibBoard.Caption = gpps2("calibboard", "window title", language$, CalibBoard.Caption)
    Label1.Caption = get_thing("calibboard", "label1", language$, Label1.Caption, Label1, default_font)
    Label2.Caption = get_thing("calibboard", "label2", language$, Label2.Caption, Label2, default_font)
    Label3.Caption = get_thing("calibboard", "label3", language$, Label3.Caption, Label3, default_font)
    set_fontstuff Label4, default_font
    set_fontstuff Label5, default_font
    bdexit.Caption = gpps2("calibboard", "bdexit", language$, bdexit.Caption)
    set_fontname bdexit, default_font
    For i = 0 To 2
        Option1(i).Caption = get_thing("calibboard", "option" + Str$(i), language$, Option1(i).Caption, Option1(i), default_font)
    Next i
    
    ' Other text
    ts$(1) = gpps2("calibboard", "ts1", language$, "Unit")
    ts$(2) = gpps2("calibboard", "ts2", language$, "Delay")
    ts$(3) = gpps2("calibboard", "ts3", language$, "Average")
    ts$(4) = gpps2("calibboard", "ts4", language$, "Span")
    ts$(5) = gpps2("calibboard", "ts5", language$, "Zero")
    ts$(6) = gpps2("calibboard", "ts6", language$, "should be")
    
End Sub
