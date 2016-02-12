VERSION 5.00
Begin VB.Form GetCompressionPressure 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Compression Pressure"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.OptionButton Option1 
         Caption         =   "Direct entry of piston pressure"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   5415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Enter compressive force on sample"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   5415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   2640
         Width           =   4935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Compression Pressure (this gets rewritten in the form load)"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Diameter of area being compressed (this gets rewritten in the form load)"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "Compressive Force on Sample (this gets rewritten in the form load)"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   2400
         Width           =   5055
      End
   End
End
Attribute VB_Name = "GetCompressionPressure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ts$(4)                              ' Text strings for this form

Private Sub Command1_Click()

Dim X As Double

If val(Text2.Text) <= 0 Then
    MsgBox ts$(1)           ' "Diameter must be greater than zero"
    Exit Sub
End If

compression_pressure = val(Text1.Text) / PCNV
If compression_pressure < 0 Then compression_pressure = 0
sample_compression_diameter = val(Text2.Text) * linear_unit_conversion#
sample_compression_pressure = val(Text3.Text) / PCNV
If sample_compression_pressure < 0 Then sample_compression_pressure = 0
If Option1.value Then
    use_sample_compression = False
Else
    use_sample_compression = True
    ' recalculate compression pressure based on above
    ' set x to cross sectional area of sample in square inches
    ' sample_compression_diameter is in cm
    X = 3.14159265 * ((sample_compression_diameter / 5.08) ^ 2)
    compression_pressure = sample_compression_pressure * X / piston_area
End If

' return value of compression_pressure
Got_Value = compression_pressure
' save the values in the user ini file
WPPS Curr_U$, "compression_pressure", Str$(compression_pressure), IFile$
WPPS Curr_U$, "sample_compression_pressure", Str$(sample_compression_pressure), IFile$
WPPS Curr_U$, "sample_compression_diameter", Str$(sample_compression_diameter), IFile$
WPPS Curr_U$, "use_sample_compression", IIf(use_sample_compression, "Y", "N"), IFile$
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

LoadTextStrings

Label1.Caption = ts$(2) + " - " + PU$           ' "Compression Pressure"
Text1.Text = Format$(compression_pressure * PCNV)
Label2.Caption = ts$(3) + " - " + linear_unit_name$     ' "Diameter of area being compressed"
Text2.Text = Format$(sample_compression_diameter / linear_unit_conversion#)
Label3.Caption = ts$(4) + " - " + PU$                   '"Compressive Force on Sample"
Text3.Text = Format$(sample_compression_pressure * PCNV)
Option1.value = Not use_sample_compression
Option2.value = use_sample_compression
'edc 12-11-06 alter border color and caption
Me.Caption = Me.Caption & "    " & SubCaption
Me.BackColor = lngBorderColor

End Sub

Public Sub LoadTextStrings()
' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    GetCompressionPressure.Caption = get_thing("getcomppress", "window title", language$, GetCompressionPressure.Caption, GetCompressionPressure, default_font)
    set_fontstuff Label1, default_font
    set_fontstuff Label2, default_font
    set_fontstuff Label3, default_font
    Option1.Caption = get_thing("getcomppress", "option1", language$, Option1.Caption, Option1, default_font)
    Option2.Caption = get_thing("getcomppress", "option2", language$, Option2.Caption, Option2, default_font)
    Command1.Caption = gpps2("getcomppress", "command1", language$, Command1.Caption)
    Command2.Caption = gpps2("getcomppress", "command2", language$, Command2.Caption)
    set_fontname Command1, default_font
    set_fontname Command2, default_font
    
    ' Other text
    ts$(1) = gpps2("getcomppress", "ts1", language$, "Diameter must be greater than zero")
    ts$(2) = gpps2("getcomppress", "ts2", language$, "Compression Pressure")
    ts$(3) = gpps2("getcomppress", "ts3", language$, "Diameter of area being compressed")
    ts$(4) = gpps2("getcomppress", "ts4", language$, "Compressive Force on Sample")
    
End Sub
