VERSION 5.00
Begin VB.Form AboutBox 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About CapWin"
   ClientHeight    =   2535
   ClientLeft      =   1410
   ClientTop       =   1920
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "ABOUTCAP.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2535
   ScaleWidth      =   5790
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.PictureBox Pic_ApplicationIcon 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         Picture         =   "ABOUTCAP.frx":030A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox Pic_HorizontalLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   30
         Left            =   975
         ScaleHeight     =   0
         ScaleWidth      =   4350
         TabIndex        =   2
         Top             =   1410
         Width           =   4380
      End
      Begin VB.CommandButton Cmd_OK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   360
         Left            =   4065
         TabIndex        =   1
         Top             =   1785
         Width           =   1275
      End
      Begin VB.Label Lbl_IconWorks 
         Caption         =   "CapWin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   975
         TabIndex        =   7
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Lbl_Version 
         Caption         =   "Version "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   6
         Top             =   810
         Width           =   2580
      End
      Begin VB.Label Lbl_Microsoft 
         Caption         =   "PMI Automated Capillary Flow Porometer for Windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   5
         Top             =   1095
         Width           =   5040
      End
      Begin VB.Label Lbl_Info 
         Caption         =   "32-bit version for Windows 95 and higher"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   990
         TabIndex        =   4
         Top             =   1530
         Width           =   4275
      End
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z
Dim ts$(2)                  ' Text strings for this form

Private Sub Cmd_OK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim temp$
    
    LoadTextStrings
    
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    'RemoveSysMenu AboutBox
    temp$ = ts$(1) + ": " + S_Version           ' "Version"
    Lbl_Version.Caption = temp$
    If H2OPERM Then
        Lbl_Microsoft.Caption = "PMI Automated Perm-Porometer for Windows"
    Else
        Lbl_Microsoft.Caption = "PMI Automated CapFlow for Windows"
    End If
    If CFAnal Then Lbl_Microsoft.Caption = "PMI Automated CFA for Windows"
    If GasPerm Then Lbl_Microsoft.Caption = "PMI Automated Gas Permeameter for Windows"
    If itester Then Lbl_Microsoft.Caption = "PMI Automated Integrity Tester for Windows"
    ' 6.71.38x begin
    If BPTester Then
        If hydrohead_exclusive Then
            Lbl_Microsoft.Caption = "PMI Automated Hydrohead Tester for Windows"
        Else
            Lbl_Microsoft.Caption = "PMI Automated Bubble Point Tester for Windows"
        End If
    End If
    If WESA_exclusive Then
        Lbl_Microsoft.Caption = "PMI Automated Envelope Surface Area Analyzer for Windows"
    ElseIf liqpermonly Then
        Lbl_Microsoft.Caption = "PMI Automated Liquid Permeameter for Windows"
    ElseIf microflowporometer Then
        Lbl_Microsoft.Caption = "PMI Automated Microflow Porometer for Windows"
    End If
    ' 6.71.38x end
    Refresh
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(2) + str$(unitnumber)           ' "Unit"
    End If
    'edc 12-11-06 alters border color and caption
    Me.BackColor = lngBorderColor
    Me.Caption = Me.Caption & "   " & SubCaption
End Sub

Public Sub LoadTextStrings()

    ' Load text elements for this form from external .ini file
    
    ' Form elements
    AboutBox.Caption = gpps2("about", "window title", language$, "About") + " CapWin"
    Cmd_OK.Caption = gpps2("about", "ok", language$, Cmd_OK.Caption)
    set_fontname Cmd_OK, default_font
    set_fontname Lbl_Version, default_font
    
    ' Other text
    ts$(1) = gpps2("about", "ts1", language$, "Version")
    ts$(2) = gpps2("about", "ts2", language$, "Unit")
    
End Sub

