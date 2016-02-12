VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EditForm 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capillary Flow Data File Editor"
   ClientHeight    =   9510
   ClientLeft      =   5160
   ClientTop       =   2535
   ClientWidth     =   11175
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
   ScaleHeight     =   634
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9255
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   10935
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4335
         Top             =   8295
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Convert LP Data"
         Filter          =   "*.cft"
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Convert LP Data"
         Height          =   570
         Left            =   1005
         TabIndex        =   47
         Top             =   8310
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         Height          =   2475
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   2415
         ScaleWidth      =   5100
         TabIndex        =   0
         Top             =   240
         Width           =   5160
      End
      Begin VB.TextBox CItem 
         Height          =   285
         Index           =   0
         Left            =   8160
         TabIndex        =   1
         Tag             =   "The Test Dates"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox CItem 
         Height          =   285
         Index           =   1
         Left            =   8160
         TabIndex        =   2
         Tag             =   "The Line 1's"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox CItem 
         Height          =   285
         Index           =   2
         Left            =   8160
         TabIndex        =   3
         Tag             =   "The Line 2's"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox CItem 
         Height          =   285
         Index           =   3
         Left            =   8160
         TabIndex        =   4
         Tag             =   "The Sample IDs"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox CItem 
         Height          =   315
         Index           =   6
         Left            =   8160
         TabIndex        =   8
         Top             =   3000
         Width           =   2205
      End
      Begin VB.TextBox CItem 
         Height          =   315
         Index           =   7
         Left            =   8160
         TabIndex        =   9
         Top             =   3360
         Width           =   2205
      End
      Begin VB.TextBox CItem 
         Height          =   315
         Index           =   9
         Left            =   8160
         TabIndex        =   10
         Top             =   3720
         Width           =   2205
      End
      Begin VB.TextBox CItem 
         Height          =   315
         Index           =   8
         Left            =   8160
         TabIndex        =   11
         Top             =   4080
         Width           =   2205
      End
      Begin VB.TextBox CItem 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   6240
         TabIndex        =   16
         Tag             =   "The Sample IDs"
         Top             =   6240
         Width           =   4095
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   6240
         Max             =   0
         Min             =   1
         TabIndex        =   17
         Top             =   6600
         Value           =   1
         Width           =   4095
      End
      Begin VB.TextBox CItem 
         Height          =   315
         Index           =   10
         Left            =   8160
         TabIndex        =   12
         Top             =   4440
         Width           =   2205
      End
      Begin VB.TextBox CItem 
         Height          =   315
         Index           =   11
         Left            =   8160
         TabIndex        =   14
         Top             =   5160
         Width           =   2205
      End
      Begin VB.TextBox CItem 
         Height          =   315
         Index           =   12
         Left            =   8160
         TabIndex        =   13
         Top             =   4800
         Width           =   2205
      End
      Begin VB.TextBox CItem 
         Height          =   315
         Index           =   13
         Left            =   8160
         TabIndex        =   15
         Top             =   5520
         Width           =   2205
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Circular"
         Height          =   255
         Index           =   0
         Left            =   8160
         TabIndex        =   5
         Top             =   2040
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Square"
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   6
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Default (0.715)"
         Height          =   255
         Left            =   7440
         TabIndex        =   18
         Top             =   7200
         Width           =   2535
      End
      Begin VB.TextBox CItem 
         Height          =   285
         Index           =   14
         Left            =   7440
         TabIndex        =   20
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton killtempcmd 
         Caption         =   "Remove Temperature Data"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   7440
         Width           =   3015
      End
      Begin VB.Label lblWetFluid 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   46
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Pict_label 
         BackColor       =   &H00FFFF80&
         Height          =   345
         Index           =   6
         Left            =   120
         TabIndex        =   43
         Top             =   5280
         Width           =   4665
      End
      Begin VB.Label Pict_label 
         BackColor       =   &H00FFFF80&
         Height          =   345
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   4800
         Width           =   4665
      End
      Begin VB.Label Pict_label 
         BackColor       =   &H00FFFF80&
         Height          =   345
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Top             =   4320
         Width           =   4665
      End
      Begin VB.Label Pict_label 
         BackColor       =   &H00FFFF80&
         Height          =   345
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Width           =   4665
      End
      Begin VB.Label Pict_label 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   38
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label Pict_label 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Height          =   225
         Index           =   1
         Left            =   840
         TabIndex        =   37
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Date"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   22
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Line 1"
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   23
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Line 2"
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   24
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Sample ID"
         Height          =   255
         Index           =   3
         Left            =   6120
         TabIndex        =   25
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Wet Fluid"
         Height          =   225
         Index           =   5
         Left            =   6240
         TabIndex        =   21
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Permeability Gas/Liquid"
         Height          =   345
         Index           =   16
         Left            =   5400
         TabIndex        =   27
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Surf. Ten. (dynes/cm)"
         Height          =   255
         Index           =   6
         Left            =   4800
         TabIndex        =   28
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Diameter of Sample (cm)"
         Height          =   225
         Index           =   7
         Left            =   4920
         TabIndex        =   29
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Thickness (cm)"
         Height          =   225
         Index           =   9
         Left            =   4920
         TabIndex        =   30
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Length of Hollow Cyl. (or 0) (cm)"
         Height          =   225
         Index           =   8
         Left            =   4200
         TabIndex        =   31
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Extended Parameters"
         Height          =   225
         Index           =   4
         Left            =   6240
         TabIndex        =   36
         Top             =   6000
         Width           =   4095
      End
      Begin VB.Label gas_or_liquid 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   7
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Chamber Diameter (or 0) (cm)"
         Height          =   225
         Index           =   10
         Left            =   4200
         TabIndex        =   32
         Top             =   4440
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Target Flow Rate (cc/min)"
         Height          =   435
         Index           =   11
         Left            =   5760
         TabIndex        =   34
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Mass (gm)"
         Height          =   225
         Index           =   12
         Left            =   5880
         TabIndex        =   33
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Density (cc/gm)"
         Height          =   345
         Index           =   13
         Left            =   5760
         TabIndex        =   35
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Pore Shape:"
         Height          =   225
         Index           =   15
         Left            =   5880
         TabIndex        =   26
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         Caption         =   "Tortuosity Factor:"
         Height          =   225
         Index           =   14
         Left            =   5160
         TabIndex        =   19
         Top             =   6960
         Width           =   3735
      End
      Begin VB.Label Pict_label 
         BackColor       =   &H00FFFF80&
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   3360
         Width           =   4665
      End
   End
   Begin VB.Menu editmenu 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu filemenu 
         Caption         =   "&Load"
         Index           =   0
      End
      Begin VB.Menu filemenu 
         Caption         =   "&Save"
         Index           =   1
      End
      Begin VB.Menu filemenu 
         Caption         =   "---------"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu filemenu 
         Caption         =   "&Import"
         Index           =   3
         Begin VB.Menu importmenu 
            Caption         =   "&Dry Curve"
            Index           =   0
         End
         Begin VB.Menu importmenu 
            Caption         =   "&Wet Curve"
            Index           =   1
         End
      End
      Begin VB.Menu filemenu 
         Caption         =   "&Export"
         Index           =   4
         Begin VB.Menu exportgaspermmenu 
            Caption         =   "&Gas Permeability File"
         End
      End
      Begin VB.Menu filemenu 
         Caption         =   "---------"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu filemenu 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
   Begin VB.Menu editmenu 
      Caption         =   "&Refresh"
      Index           =   2
   End
   Begin VB.Menu editmenu 
      Caption         =   "&Help"
      Index           =   3
   End
End
Attribute VB_Name = "EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim cftfile As tcftfile
Dim last_value As Integer
Dim edit_file_name As String

Dim NB_WetI As Integer
Dim MinP As Single
Dim MinQ As Single
Dim MinPD  As Single
Dim MinQD As Single
Dim MinPW As Single
Dim MinQW As Single
Dim MaxP As Single
Dim MaxQ As Single
Dim MaxPD As Single
Dim MaxQD As Single
Dim MaxPW As Single
Dim MaxQW As Single
Dim iswet As Boolean
Dim isdry As Boolean

Dim TempWetI As Integer
Dim TempDryI As Integer
Dim DelQ As Single
Dim dirty As Boolean
Dim ts$(52)                     ' Text strings for this form

Private Sub Check1_Click()

If Check1.value = 0 Then
    Check1.Caption = ts$(1) + " (0.715)"    ' "Default"
    CItem(14).Visible = False
Else
    Check1.Caption = ts$(2)                 ' "Group Defined"
    CItem(14).Visible = True
End If
dirty = True

End Sub

Private Sub CItem_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index > 5 Then
        Select Case KeyAscii
            Case 8, Asc(Decimal_Point$), 48 To 57
                dirty = True
            Case Else
                KeyAscii = 0
        End Select
    Else
        dirty = True
    End If

End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
Dim poo As String
poo = CommonDialog1.filename

    If Len(poo) > 0 Then
        Shell (App.path & "\lpc.exe" & " " & Chr$(34) & poo & Chr$(34))
    End If

End Sub

Private Sub Editmenu_Click(Index As Integer)

    Dim r As Long
    
    Select Case Index
        Case 1
        Case 2
            ' by clicking on refresh, you redraw the curves and this
            ' causes the min and max values to be reset back to full size
            MousePointer = 11
            ShowCurves
            Picture1.Tag = ""
            MousePointer = 0
        Case 3
            r = WinHelp(hwnd, HelpFile$, Help_Context, ByVal 35&)
    End Select

End Sub

Private Sub exportgaspermmenu_Click()
    ExportDry
End Sub

Private Sub filemenu_Click(Index As Integer)

Select Case Index
    Case 0 ' load
        LoadFile
        killtempcmd.Enabled = (cftfile.temperaturesensor > 0)
    Case 1 ' save
        SaveFile
    Case 6 ' exit
        Unload Me
End Select

End Sub

Private Sub Form_Load()

    Dim i As Integer
    
    LoadTextStrings
    
    ' Following is superseded by "LoadTextStrings"
    'For i = 0 To 14
    '    CItem(i).FontName = system_font
    '    CItem(i).fontsize = font_size
    '    CItem(i).fontbold = font_bold
    'Next i
    'gas_or_liquid.FontName = system_font
    'gas_or_liquid.fontsize = font_size
    'gas_or_liquid.fontbold = font_bold
    dirty = False
    TitleScrn.modifymenu(3).Enabled = False
    If GasPerm Then
        Caption = ts$(3) + " " + ts$(8)         ' "Gas Permeameter Data File Editor"
    ElseIf CFAnal Then
        Caption = ts$(4) + " " + ts$(8)         ' "CFA Data File Editor"
    ElseIf itester Then
        Caption = ts$(5) + " " + ts$(8)         ' "Integrity Tester Data File Editor"
    ElseIf BPTester Then
        ' 6.71.38x begin
        If hydrohead_exclusive Then
            Caption = ts$(51)                   ' "Hydrohead Tester Data File Editor"
        Else
            Caption = ts$(6) + " " + ts$(8)         ' "Bubble Point Tester Data File Editor"
        End If
        ' 6.71.38x end
    Else
        Caption = ts$(7) + " " + ts$(8)         ' "Capillary Flow Data File Editor"
    End If
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    initialize_cftfile cftfile
    update_screen
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(9) + str$(unitnumber)   ' "Unit"
    End If
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
If Dir(App.path & "\lpc.exe", vbNormal) <> "" Then Command1.Visible = True

End Sub

Private Sub update_screen()

Dim i As Integer

' update the screen based on current type of file
With cftfile

If .cfttype = ect_undefined Then
    filemenu(1).Enabled = False ' turn off save
    filemenu(3).Enabled = False ' turn off import
    filemenu(4).Enabled = False ' turn off export
    iswet = False
    isdry = False
    For i = 0 To 14
        If i <> 5 Then CItem(i).Text = ""
    Next i
    For i = 0 To 4
        Label1(i).Enabled = False
        CItem(i).Enabled = False
    Next i
    For i = 5 To 16
        Label1(i).Visible = False
    Next i
    For i = 6 To 14
        CItem(i).Visible = False
    Next i
    gas_or_liquid.Visible = False
    Check1.Visible = False
    Option1(0).Visible = False
    Option1(1).Visible = False
    HScroll1.Enabled = False
    Me.Caption = ts$(8) + " - " + ts$(10)           ' "Data Editor - no file loaded"
Else
    Me.Caption = ts$(8) + " - " + edit_file_name$   '"Data Editor:"
    filemenu(1).Enabled = True ' turn on save
    ' import and export only work for porometry
    ' and only porometry has a visible wet curve
    If .cfttype = ect_porometry Then
        filemenu(3).Enabled = True
        filemenu(4).Enabled = True
        iswet = True
        ' set up min and max pw and qw - don't assume that the data is perfect
        ' though it should be in order anyway
        MinQW = .wet_data.a(1).fl
        MaxQW = MinQW
        MinPW = .wet_data.a(1).pr
        MaxPW = MinPW
        For i = 2 To .num_wet_points
            If .wet_data.a(i).fl < MinQW Then
                MinQW = .wet_data.a(i).fl
            End If
            If .wet_data.a(i).fl > MaxQW Then
                MaxQW = .wet_data.a(i).fl
            End If
            If .wet_data.a(i).pr < MinPW Then
                MinPW = .wet_data.a(i).pr
            End If
            If .wet_data.a(i).pr > MaxPW Then
                MaxPW = .wet_data.a(i).pr
            End If
        Next i
    Else
        iswet = False
        filemenu(3).Enabled = False
        filemenu(4).Enabled = False
    End If
    ' only porometry, permeability, esa, and efd have dry curves
    If .cfttype = ect_porometry Or .cfttype = ect_esa Or .cfttype = ect_efd _
        Or .cfttype = ect_permeability Then
        isdry = True
        ' set up min and max pd and qd - don't assume that the data is perfect
        ' though it should be in order anyway
        MinQD = .dry_data.a(1).fl
        MaxQD = MinQD
        MinPD = .dry_data.a(1).pr
        MaxPD = MinPD
        For i = 2 To .num_dry_points
            If .dry_data.a(i).fl < MinQD Then
                MinQD = .dry_data.a(i).fl
            End If
            If .dry_data.a(i).fl > MaxQD Then
                MaxQD = .dry_data.a(i).fl
            End If
            If .dry_data.a(i).pr < MinPD Then
                MinPD = .dry_data.a(i).pr
            End If
            If .dry_data.a(i).pr > MaxPD Then
                MaxPD = .dry_data.a(i).pr
            End If
        Next i
    Else
        isdry = False
    End If
    ' first, update stuff for all data file types
    For i = 0 To 3
        Label1(i).Enabled = True
        CItem(i).Enabled = True
    Next i
    CItem(0).Text = .date
    CItem(1).Text = .Line1
    CItem(2).Text = .Line2
    CItem(3).Text = .sample_id
    If .num_extended > 0 Then
        Label1(4).Enabled = True
        CItem(4).Enabled = True
        CItem(4).Text = .extended_value(1)
        last_value = 0 ' so that when we change the value we don't update anything
        HScroll1.max = .num_extended
        HScroll1.value = 1 ' this will change last_value to 1
        HScroll1.Enabled = True
    Else
        Label1(4).Enabled = False
        CItem(4).Enabled = False
        CItem(4).Text = ""
        HScroll1.Enabled = False
        last_value = 0
    End If
    If .cfttype = ect_efd Or .cfttype = ect_esa Or .cfttype = ect_permeability Or .cfttype = ect_microflow Then
        Label1(16).Visible = True
        gas_or_liquid.Visible = True
        gas_or_liquid.Enabled = True
        Label1(5).Visible = False
        lblWetFluid.Visible = False 'CItem(5).Visible = False rem edc 3-21-07
        Label1(6).Visible = False
        CItem(6).Visible = False
        If .pgl.type = epgl_other Then
            gas_or_liquid = .pgl.name + " mu=" + str$(val(.pgl.viscosity))
        Else
            gas_or_liquid = .pgl.name
        End If
    ElseIf .cfttype = ect_resin Then
        ' resin test has fluid name, but it is displayed as a liquid name and doesn't have a surface tension
        ' also, this fliuid name can not be edited (it crashes the system if you try)
        Label1(16).Visible = True
        gas_or_liquid.Visible = True
        gas_or_liquid.Caption = .fluidname
        gas_or_liquid.Enabled = False
        Label1(5).Visible = False
        lblWetFluid.Visible = False 'CItem(5).Visible = False rem edc 3-21-07
        Label1(6).Visible = False
        CItem(6).Visible = False
    Else
        Label1(16).Visible = False
        gas_or_liquid.Visible = False
        lblWetFluid.Visible = True    'CItem(5).Visible = True rem edc 3-21-07
        lblWetFluid.Caption = .fluidname 'CItem(5).Text = .fluidname   remedc03-21-07
        Label1(5).Visible = True
        If .cfttype = ect_phold Or .cfttype = ect_microflow Then
            Label1(5).Caption = ts$(11)         ' "Gas Name"
            Label1(6).Visible = False
            CItem(6).Visible = False
        Else
            Label1(5).Caption = ts$(12)         ' "Fluid Name"
            Label1(6).Visible = True
            CItem(6).Visible = True
            CItem(6).Text = str$(.surfTen)
        End If
    End If
    ' phold and hydrohead don't have diameter (or cyl len or thickness)
    If .cfttype = ect_phold Or .cfttype = ect_hydrohead Then
        For i = 7 To 9
            Label1(i).Visible = False
            CItem(i).Visible = False
        Next i
    Else
        ' for all others, sometimes the diameter is not enabled
        ' in which case it should be set to 0
        If Not .diam_valid Then
            .diameter = 0
            .cyl_len = 0
            .thickness = 0
        End If
        For i = 7 To 9
            Label1(i).Visible = True
            CItem(i).Visible = True
        Next i
        CItem(7).Text = Format$(.diameter)
        CItem(8).Text = Format$(.cyl_len)
        CItem(9).Text = Format$(.thickness)
    End If
    ' stuff only for esa and efd:
    If .cfttype = ect_efd Or .cfttype = ect_esa Then
        Label1(10).Visible = True
        CItem(10).Visible = True
        If .scdiam_valid Then
            CItem(10).Text = Format$(.scdiameter)
        Else
            CItem(10).Text = "0"
        End If
        ' 11 is only for efdflow
        If .cfttype = ect_efd Then
            Label1(11).Visible = True
            Label1(11).Caption = ts$(13)        ' "Target Flow Rate (cc/min)"
            CItem(11).Visible = True
            CItem(11).Text = Format$(.efd_flow)
        Else
            Label1(11).Visible = False
            CItem(11).Visible = False
        End If
        Label1(12).Visible = True
        CItem(12).Visible = True
        CItem(12).Text = Format$(.mass)
        Label1(13).Visible = True
        CItem(13).Visible = True
        CItem(13).Text = Format$(.density)
    Else
        For i = 10 To 13
            Label1(i).Visible = False
            CItem(i).Visible = False
        Next i
    End If
    If .cfttype = ect_microflow Then
        Label1(11).Visible = True
        Label1(11).Caption = ts$(14)            ' "Chamber Volume (cc)"
        CItem(11).Visible = True
        CItem(11).Text = Format$(.chamber_volume)
    End If
    ' finally, for stuff only using wetting fluids:
    If .cfttype = ect_bubblepoint Or .cfttype = ect_porometry Or .cfttype = ect_integrity Then
        ' tortuosity factor
        Label1(14).Visible = True
        Check1.Visible = True
        ' setting the value of the check box will automatically cause
        ' the caption of the check box to change and for the text
        ' box associated with it to become visible or not.
        If .tfactor = 0.715 Then
            Check1.value = 0
        Else
            Check1.value = 1
            CItem(14).Text = Format$(.tfactor)
        End If
        ' square pores - only if enabled in the software or if the data
        ' file is set with square pores
        ' default to circular
        Option1(0).value = True
        If (Not SqrPore) And .square_pores = False Then
            Option1(0).Visible = False
            Option1(1).Visible = False
            Label1(15).Visible = False
        Else
            Option1(0).Visible = True
            Option1(1).Visible = True
            Label1(15).Visible = True
            Option1(1).value = .square_pores
        End If
    Else
        Label1(14).Visible = False
        Label1(15).Visible = False
        Check1.Visible = False
        CItem(14).Visible = False
        Option1(0).Visible = False
        Option1(1).Visible = False
    End If
End If
End With
ShowCurves

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pict_label(2).Caption = ""
    Pict_label(1).Caption = ""
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)

If dirty Then
    If MsgBox(ts$(15), vbYesNo) = vbNo Then     ' "Do you want to close the editor without saving your changes?"
        cancel = 1
    End If
End If

End Sub

Private Sub Form_Unload(cancel As Integer)
    TitleScrn.modifymenu(3).Enabled = True
End Sub

Private Sub gas_or_liquid_Click()

selected_gas_or_liquid = cftfile.pgl
NeedLiquid = True
NeedGas = True
GLSel.Show 1
If selected_gas_or_liquid.type <> epgl_undefined Then
    cftfile.pgl = selected_gas_or_liquid
    dirty = True
End If
With selected_gas_or_liquid
    If .type <> epgl_other Then
        gas_or_liquid.Caption = .name
        CItem(6).Text = selected_gas_or_liquid.surfTen
        dirty = True
    Else
        gas_or_liquid.Caption = .name
        CItem(6).Text = selected_gas_or_liquid.surfTen
        dirty = True
    End If
End With
End Sub
Private Sub lblWetFluid_Click()
Dim strTemp As String
strTemp = cftfile.fluidname
If strTemp = "GALWICK" Then
    selected_gas_or_liquid.name = "Galwick"
    ElseIf strTemp = "POREWICK" Then
    selected_gas_or_liquid.name = "Porewick"
    ElseIf strTemp = "SILWICK" Then
    selected_gas_or_liquid.name = "Silwick"
    Else
selected_gas_or_liquid = cftfile.pgl
End If
NeedFluid = True
NeedGas = False
NeedLiquid = False
GLSel.Show 1
If selected_gas_or_liquid.type <> epgl_other Then
    cftfile.pgl = selected_gas_or_liquid
    dirty = True
End If
With selected_gas_or_liquid
   If .type <> epgl_other Then
        lblWetFluid.Caption = .name
        CItem(6).Text = selected_gas_or_liquid.surfTen
        dirty = True
    Else
        lblWetFluid.Caption = .name
        CItem(6).Text = CStr(selected_gas_or_liquid.surfTen)
        dirty = True
    End If
End With
End Sub


Private Sub HScroll1_Change()
    
    Dim temp$
    
    If last_value > 0 Then
        cftfile.extended_value(last_value) = CItem(4).Text
    End If
    
    ' Trap English output and divert it to a translatable text string -
    ' works for now, but only for specific strings!
    temp$ = LCase$(cftfile.extended_name(HScroll1.value))
    If temp$ = "operator" Then
        temp$ = ts$(43)         ' "Operator"
    ElseIf temp$ = "lot number" Then
        temp$ = ts$(44)         ' "Lot Number"
    ElseIf temp$ = "hardware serial number" Then
        temp$ = ts$(45)         ' "Hardware Serial Number"
    ElseIf temp$ = "type of test" Then
        temp$ = ts$(46)         ' "Type of Test"
    ElseIf temp$ = "wet parameter" Then
        temp$ = ts$(47)         ' "Wet Parameter"
    ElseIf temp$ = "dry parameter" Then
        temp$ = ts$(48)         ' "Dry Parameter"
    Else
        temp$ = cftfile.extended_name(HScroll1.value)
    End If
    Label1(4).Caption = temp$
    
    CItem(4).Text = cftfile.extended_value(HScroll1.value)
    last_value = HScroll1.value

End Sub

Private Sub LoadFile()
    
    If dirty Then
        If MsgBox(ts$(50), vbOKCancel) = vbCancel Then         ' "WARNING: You will lose any changes you may have made to the current file."
            Exit Sub
        End If
    End If
    fsel_name$ = ""
    If GasPerm Then
        fsel_title$ = ts$(3) + " " + ts$(17)    ' "Gas Permeameter Data File."
    ElseIf CFAnal Then
        fsel_title$ = ts$(4) + " " + ts$(17)    ' "CFA Data File."
    ElseIf itester Then
        fsel_title$ = ts$(5) + " " + ts$(17)    ' "Integrity Tester Data File"
    ElseIf BPTester Then
        ' 6.71.38x begin
        If hydrohead_exclusive Then
            fsel_title$ = ts$(52)               ' "Hydrohead Tester Data File"
        Else
            fsel_title$ = ts$(6) + " " + ts$(17)    ' "Bubble Point Tester Data File"
        End If
        ' 6.71.38x end
    Else
        fsel_title$ = ts$(7) + " " + ts$(17)    ' "Capillary Flow Data File."
    End If
    fsel_path$ = path(0)
    fsel_io = True
    fsel Me.hwnd
    If fsel_return$ = "" Then
        Exit Sub
    End If
    path(0) = fsel_path$
    edit_file_name$ = fsel_return$
    load_cftfile cftfile, edit_file_name$
    NB_WetI = 1
    TempWetI = cftfile.num_wet_points
    TempDryI = cftfile.num_dry_points
    update_screen
    dirty = False

End Sub

Private Sub ImportDry()

    Dim X As tcftfile, i As Integer
    
    fsel_name$ = ""
    fsel_title$ = ts$(18) + " " + ts$(21)             ' "Import Dry Curve"
    fsel_path$ = path(0)
    fsel_io = True
    fsel Me.hwnd
    If fsel_return$ = "" Then
        Exit Sub
    End If
    path(0) = fsel_path$
    load_cftfile X, fsel_return$
    ' file must be a porometry, gasperm, esa, or efd file
    If X.cfttype <> ect_efd And X.cfttype <> ect_esa And X.cfttype <> ect_permeability _
        And X.cfttype <> ect_porometry Then
        MsgBox ts$(19) + " " + ts$(21) + " - " + ts$(20)             ' "Error:  This data file does not contain a dry curve - import cancelled"
        Exit Sub
    End If
    With cftfile
    ' put the dry curve from the imported file into the current data structure
    .dry_data = X.dry_data
    .num_dry_points = X.num_dry_points
    ' rescale from the imported file's p0 and q0 values to the current data
    ' structure's p0 and q0 values
    For i = 1 To .dry_data.s
        .dry_data.a(i).fl = .dry_data.a(i).fl - X.Q0 + .Q0
        .dry_data.a(i).pr = .dry_data.a(i).pr - X.P0 + .P0
    Next i
    TempDryI = .num_dry_points
    End With
    ' refresh the screen display
    update_screen
    dirty = True

End Sub

Private Sub ImportWet()

    Dim X As tcftfile, i As Integer
    
    fsel_name$ = ""
    fsel_title$ = ts$(18) + " " + ts$(22)   ' "Import Wet Curve"
    fsel_path$ = path(0)
    fsel_io = True
    fsel Me.hwnd
    If fsel_return$ = "" Then
        Exit Sub
    End If
    path(0) = fsel_path$
    load_cftfile X, fsel_return$
    ' file must be a porometry or integrity file
    ' though using an integrity file may be strange
    If X.cfttype <> ect_porometry And X.cfttype <> ect_integrity Then
        MsgBox ts$(19) + " " + ts$(22) + " - " + ts$(20) ' "Error:  This data file does not contain a wet curve - import cancelled"
        Exit Sub
    End If
    With cftfile
    ' put the wet curve from the imported file into the current data structure
    .wet_data = X.wet_data
    .num_wet_points = X.num_wet_points
    ' use the p0 and q0 values of the imported file
    ' rescale the dry curve of the current file to compensate
    For i = 1 To .dry_data.s
        .dry_data.a(i).fl = .dry_data.a(i).fl - .Q0 + X.Q0
        .dry_data.a(i).pr = .dry_data.a(i).pr - .P0 + X.P0
    Next i
    .Q0 = X.Q0
    .P0 = X.P0
    ' also use the fluidname and surface tension from the imported file
    ' since the wet curve depends on this
    .fluidname = X.fluidname
    .surfTen = X.surfTen
    TempWetI = .num_wet_points
    End With
    ' refresh the screen display
    NB_WetI = 1
    update_screen
    dirty = True

End Sub

Private Sub importmenu_Click(Index As Integer)

If Index = 0 Then
    ImportDry
Else
    ImportWet
End If

End Sub

Private Sub killtempcmd_Click()
killtempcmd.Enabled = False
cftfile.temperaturesensor = 0
End Sub



Private Sub Picture1_KeyPress(KeyAscii As Integer)

    Dim tempx$, i%
    Dim LX As Single
    Dim LY As Single
    
    On Error GoTo Ed_Error
    If Pict_label(0).Tag = "dis" Then Exit Sub
    Pict_label(0).Tag = "dis"
    Pict_label(1).Caption = ""
    Pict_label(2).Caption = ""

    tempx$ = UCase$(Chr$(KeyAscii))
    If (tempx$ = "B") And iswet Then
        LX = cftfile.wet_data.a(1).pr
        LY = cftfile.wet_data.a(1).fl
        i% = 1
        ' find first point in wet curve whose pressure is >= the current
        ' mouse pointer (MaxP)
        While LX < MaxP And i% < cftfile.num_wet_points
            i% = i% + 1
            ' blacken out the line to the next data point
            Picture1.Line (LX, LY)-(cftfile.wet_data.a(i%).pr, _
                cftfile.wet_data.a(i%).fl), RGB(255, 255, 255)
            LX = cftfile.wet_data.a(i%).pr
            LY = cftfile.wet_data.a(i%).fl
        Wend
        ' NB_WetI = the index of the new first wet curve data point
        NB_WetI = i%
        If NB_WetI = cftfile.num_wet_points Then
            NB_WetI = NB_WetI - 1
        End If
        dirty = True
    End If

    If (tempx$ = "R") Then
        DelQ = 1 ' set the dry curve multiplier to 1
        NB_WetI = 1 ' reset the bubble point
        ' refresh the screen
        Editmenu_Click 2
    End If

    If tempx$ = "M" Then
        ' move the dry curve so that it hits the current point of the mouse
        ' note that the dry data is now in ascending order.  In previous
        ' versions of this program, the dry curve was kept in descending
        ' order
        i% = 1
        While cftfile.dry_data.a(i%).pr <= MaxP And i% < cftfile.num_dry_points
            i% = i% + 1
        Wend
        DelQ = MaxQ / cftfile.dry_data.a(i%).fl
        Picture1.Tag = "move"
        Editmenu_Click 2
        dirty = True
    End If

    If (tempx$ = "Z") Or (tempx$ = "T") Then
        ' zoom in or trim
        If tempx$ = "T" Then Picture1.Tag = "trim"
        Picture1.MousePointer = 11
        On Error Resume Next
        Picture1.Cls
        Picture1.Scale (MinP, MaxQ * DelQ)-(MaxP, MinQ)
        If iswet Then
            LX = cftfile.wet_data.a(1).pr
            LY = cftfile.wet_data.a(1).fl
            i% = 1
            While (LX <= MaxP) And (LY <= MaxQ) And (i% < cftfile.num_wet_points)
                i% = i% + 1
                If i% > NB_WetI Then
                    Picture1.Line (LX, LY)-(cftfile.wet_data.a(i%).pr, _
                        cftfile.wet_data.a(i%).fl), RGB(0, 0, 255)
                End If
                LX = cftfile.wet_data.a(i%).pr
                LY = cftfile.wet_data.a(i%).fl
            Wend
            TempWetI = i%
            Picture1.MousePointer = 2
        End If
        If isdry Then
            Picture1.MousePointer = 11
            LX = cftfile.dry_data.a(1).pr
            LY = cftfile.dry_data.a(1).fl * DelQ
            i% = 1
            While (LX <= MaxP) And (LY <= MaxQ) And (i% < cftfile.num_dry_points)
                i% = i% + 1
                Picture1.Line (LX, LY)-(cftfile.dry_data.a(i%).pr, _
                    cftfile.dry_data.a(i%).fl * DelQ), RGB(255, 0, 0)
                LX = cftfile.dry_data.a(i%).pr
                LY = cftfile.dry_data.a(i%).fl * DelQ
            Wend
            TempDryI = i%
        End If
        On Error GoTo 0
        Picture1.MousePointer = 2
        dirty = True
    End If
    Exit Sub
Ed_Error:
    MsgBox ts$(23), 16, ts$(24)    ' "This edit operation is not possible."/"Illegal Operation"
    Resume Ed_Finis
Ed_Finis:
    On Error GoTo 0

End Sub

Private Sub Picture1_LostFocus()
    Pict_label(2).Caption = ""
    Pict_label(1).Caption = ""
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' this shows the differential pressure when you click the button
    If (Not iswet) And (Not isdry) Then Exit Sub
    Picture1.SetFocus
    Pict_label(0).Tag = ""
    Pict_label(1).Caption = "P = " + Format$((X - Picture1.ScaleLeft) * PCNV, "##0.000 ") + PU$
    Pict_label(2).Caption = "F = " + Format$(Y) + " cc/min"
    ' it also updates the maxp and maxq variables which are used when
    ' you press a button to trim, move, etc..
    MaxP = X
    MaxQ = Y
    ' this doesn't do much of anything
    Refresh

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' this shows the absolute pressure when you just move the mouse
    If (Not iswet) And (Not isdry) Then Exit Sub
    Picture1.SetFocus
    Pict_label(0).Tag = ""
    Pict_label(1).Caption = "P = " + Format$(X * PCNV, "##0.000 ") + PU$
    Pict_label(2).Caption = "F = " + Format$(Y + MinQ) + " cc/min"
    ' it also updates the maxp and maxq variables which are used when
    ' you press a button to trim, move, etc..
    MaxP = X
    MaxQ = Y
    ' this doesn't do much of anything
    Refresh

End Sub

Private Function check_values() As Boolean

Dim X As Single

' returns true if there are any errors in the information on the screen
' updates the cftfile information at the same time, so call this
' right before you save the file
With cftfile
' first, make sure that the extended information is saved, if any
If last_value > 0 Then
    .extended_value(last_value) = CItem(4).Text
End If
' next, things that are always valid
.date = CItem(0).Text
.Line1 = CItem(1).Text
.Line2 = CItem(2).Text
.sample_id = CItem(3).Text
If .cfttype = ect_efd Or .cfttype = ect_esa Or .cfttype = ect_permeability Then
    ' they use the perm gas/liquid, and that is already set
    ' but we will check it just in case
    If .pgl.type = epgl_undefined Then
        MsgBox ts$(26)                  ' "Error: permeability gas or liquid type not set properly"
        check_values = True
        GoTo done
    End If
ElseIf .cfttype = ect_resin Then
    ' resin test stores liquid name in .fluidname, but edits it as if it was a permeability liquid name
    ' this may change in the future
    .fluidname = gas_or_liquid.Caption
Else
    ' everyone else uses the fluid name, which can be anything
    .fluidname = lblWetFluid.Caption            'CItem(5).Text remedc 3-22-07
    ' and everyone but phold and microflow and resin have a surface tension
    If .cfttype <> ect_phold And .cfttype <> ect_microflow Then
        X = myVal(CItem(6).Text)
        If X <= 0 Then
            MsgBox ts$(27)              ' "Error: Surface Tension incorrect"
            check_values = True
            CItem(6).SetFocus
            GoTo done
        End If
        .surfTen = X
    End If
End If
' pressure hold and hydrohead don't have dimensions
If .cfttype <> ect_phold And .cfttype <> ect_hydrohead Then
    ' diameter determines if other things are valid as well
    X = myVal(CItem(7).Text)
    If X = 0 Then
        ' for porometry types, diameter is optional
        If .cfttype = ect_bubblepoint Or .cfttype = ect_integrity Or _
           .cfttype = ect_porometry Then
           .diam_valid = False
        Else
            ' for other types, it is not valid
            MsgBox ts$(28)          ' "Error: Invalid diameter"
            check_values = True
            CItem(7).SetFocus
            GoTo done
        End If
    ElseIf X < 0 Then
        ' negative is never valid
        MsgBox ts$(28)              ' "Error: Invalid diameter"
        check_values = True
        CItem(7).SetFocus
        GoTo done
    Else    ' valid diameter - check for other stuff
        .diam_valid = True
        .diameter = X
        X = myVal(CItem(8).Text)
        ' 0 is valid to signify no cyl. len
        ' negative is not valid
        If X < 0 Then
            MsgBox ts$(29)          ' "Error: Invalid Cyl. Len."
            check_values = True
            CItem(8).SetFocus
            GoTo done
        End If
        .cyl_len = X
        X = myVal(CItem(9).Text)
        ' thickness must be greater than 0
        If X <= 0 Then
            MsgBox ts$(30)          ' "Error: Invalid Thickness"
            check_values = True
            CItem(9).SetFocus
            GoTo done
        End If
        .thickness = X
    End If
End If
' stuff only in esa and efd
If .cfttype = ect_efd Or .cfttype = ect_esa Then
    X = myVal(CItem(10).Text)
    ' 0 signifies no sc diameter
    If X < 0 Then
        MsgBox ts$(31)              ' "Error: Invalid SC Diameter"
        check_values = True
        CItem(10).SetFocus
        GoTo done
    End If
    If X = 0 Then
        .scdiam_valid = False
    Else
        .scdiam_valid = True
        .scdiameter = X
    End If
    ' mass
    X = myVal(CItem(12).Text)
    If X <= 0 Then
        MsgBox ts$(32)              ' "Error: Invalid mass"
        check_values = True
        CItem(12).SetFocus
        GoTo done
    End If
    .mass = X
    ' density
    X = myVal(CItem(13).Text)
    If X <= 0 Then
        MsgBox ts$(33)              ' "Error: Invalid density"
        check_values = True
        CItem(13).SetFocus
        GoTo done
    End If
    .density = X
End If
' efdflow only in efd
If .cfttype = ect_efd Then
    X = myVal(CItem(11).Text)
    ' must be positive
    If X <= 0 Then
        MsgBox ts$(34)              ' "Error: Invalid efd flow"
        check_values = True
        CItem(11).SetFocus
        GoTo done
    End If
    .efd_flow = X
End If
' chamber volume only in microflow (uses same spot as eft flow)
If .cfttype = ect_microflow Then
    X = myVal(CItem(11).Text)
    ' must be positive
    If X <= 0 Then
        MsgBox ts$(35)              ' "Error: Invalid chamber volume"
        check_values = True
        CItem(11).SetFocus
        GoTo done
    End If
    .chamber_volume = X
End If
' pore stuff only porometry and related set tests
If .cfttype = ect_bubblepoint Or .cfttype = ect_integrity _
    Or .cfttype = ect_porometry Then
    ' tortuosity factor
    If Check1.value = 0 Then
        .tfactor = 0.715
    Else
        X = myVal(CItem(14).Text)
        If X <= 0 Then
            MsgBox ts$(36)          ' "Error: Invalid tortuosity factor"
            check_values = True
            CItem(14).SetFocus
            GoTo done
        End If
        .tfactor = X
    End If
    ' square pores
    .square_pores = Option1(1).value
End If
' returns false if everything is ok to save
check_values = False
done:
End With

End Function

Private Sub SaveFile()

    Dim X As tcftfile, i As Integer
    
    If check_values Then Exit Sub
    If GasPerm Then
        fsel_title$ = ts$(3) + " " + ts$(17)    ' "Gas Permeameter Data File."
    ElseIf CFAnal Then
        fsel_title$ = ts$(4) + " " + ts$(17)    ' "CFA Data File."
    ElseIf itester Then
        fsel_title$ = ts$(5) + " " + ts$(17)    ' "Integrity Tester Data File."
    ElseIf BPTester Then
        ' 6.71.38x begin
        If hydrohead_exclusive Then
            fsel_title$ = ts$(52)               ' "Hydrohead Tester Data File"
        Else
            fsel_title$ = ts$(6) + " " + ts$(17)    ' "Bubble Point Tester Data File."
        End If
        ' 6.71.38x end
    Else
        fsel_title$ = ts$(7) + " " + ts$(17)    ' "Capillary Flow Data File."
    End If
    fsel_name$ = edit_file_name$
    fsel_path$ = path(0)
    fsel_io = False
    fsel Me.hwnd
    If fsel_return$ = "" Then
        Exit Sub
    End If
    path(0) = fsel_path$
    edit_file_name$ = fsel_return$
    X = cftfile
    If Picture1.Tag = "trim" Then
        X.num_dry_points = TempDryI
        X.num_wet_points = TempWetI
    End If
    If NB_WetI > 1 And X.num_wet_points > NB_WetI Then
        X.num_wet_points = X.num_wet_points - NB_WetI + 1
        For i = 1 To X.num_wet_points
            X.wet_data.a(i) = X.wet_data.a(i + NB_WetI - 1)
        Next i
    End If
    If DelQ <> 1 Then
        For i = 1 To X.num_dry_points
            X.dry_data.a(i).fl = X.dry_data.a(i).fl * DelQ
        Next i
    End If
    save_cftfile X, edit_file_name$
    dirty = False
    'update_screen

End Sub

Private Sub ExportDry()

    Dim X As tcftfile, i As Integer
    
    If check_values Then Exit Sub
    If Not cftfile.diam_valid Then
        MsgBox ts$(37)                  ' "Please supply the sample diameter and thickness"
        CItem(7).SetFocus
        Exit Sub
    End If
    X = cftfile ' make a temporary copy
    If X.cfttype = ect_porometry Then
        ' force it to gas type "Air"
        X.pgl.type = epgl_air
        complete_pgl X.pgl
    End If
    If DelQ <> 1 Then
        For i = 1 To X.num_dry_points
            X.dry_data.a(i).fl = X.dry_data.a(i).fl * DelQ
        Next i
    End If
    X.cfttype = ect_permeability
    fsel_name$ = ""
    If GasPerm Then
        fsel_title$ = ts$(3) + " " + ts$(17)        ' "Gas Permeameter Data File."
    ElseIf CFAnal Then
        fsel_title$ = ts$(4) + " " + ts$(17)        ' "CFA Data File."
    ElseIf itester Then
        fsel_title$ = ts$(5) + " " + ts$(17)        ' "Integrity Tester Data File."
    Else
        fsel_title$ = ts$(6) + " " + ts$(17)        ' "Capillary Flow Data File."
    End If
    fsel_path$ = path(0)
    fsel_io = False
    fsel Me.hwnd
    If fsel_return$ = "" Then
        Exit Sub
    End If
    path(0) = fsel_path$
    save_cftfile X, fsel_return$

End Sub

Private Sub ShowCurves()

    Dim tempx$, i%
    Dim LX As Single
    Dim LY As Single

If Picture1.Tag <> "move" Then
    DelQ = 1
Else
    Picture1.Tag = ""
End If
Picture1.Cls
If (Not iswet) And (Not isdry) Then Exit Sub
If iswet Then
    tempx$ = ts$(38)        ' "(B)ubble Pt."
Else
    tempx$ = ""
End If
Pict_label(0).Caption = tempx$
Pict_label(3).Caption = ts$(39)     ' "(T)rim"
Pict_label(4).Caption = ts$(40)     '"(Z)oom"
Pict_label(5).Caption = ts$(41)     '"(R)estore"
Pict_label(6).Caption = ts$(42)     ' "(M)ove Dry"
If iswet Then
    MinQ = MinQW: MaxQ = MaxQW
    MinP = MinPW: MaxP = MaxPW
Else
    MinQ = MinQD: MaxQ = MaxQD
    MinP = MinPD: MaxP = MaxPD
End If
If isdry Then
    If MinQ > MinQD Then MinQ = MinQD
    If MinP > MinPD Then MinP = MinPD
    If MaxQ < MaxQD Then MaxQ = MaxQD
    If MaxP < MaxPD Then MaxP = MaxPD
End If
If MinQ = MaxQ Then MaxQ = MaxQ + 1
If MinP = MaxP Then MaxP = MaxP + 1

Picture1.Scale (MinP, MaxQ * DelQ)-(MaxP, MinQ)

With cftfile
If iswet Then
    LX = .wet_data.a(NB_WetI).pr
    LY = .wet_data.a(NB_WetI).fl
    For i% = NB_WetI + 1 To .num_wet_points
        Picture1.Line (LX, LY)-(.wet_data.a(i%).pr, .wet_data.a(i%).fl), RGB(0, 0, 255)
        LX = .wet_data.a(i%).pr
        LY = .wet_data.a(i%).fl
    Next i%
End If

If isdry Then
    LX = .dry_data.a(1).pr
    LY = .dry_data.a(1).fl * DelQ
    For i% = 2 To .num_dry_points
        Picture1.Line (LX, LY)-(.dry_data.a(i%).pr, .dry_data.a(i%).fl * DelQ), RGB(255, 0, 0)
        LX = .dry_data.a(i%).pr
        LY = .dry_data.a(i%).fl * DelQ
    Next i%
End If
End With

End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    
    ' Window caption is set dynamically
    editmenu(1).Caption = gpps2("edit", "editmenu1", language$, editmenu(1).Caption)
        filemenu(0).Caption = gpps2("edit", "filemenu0", language$, filemenu(0).Caption)
        filemenu(1).Caption = gpps2("edit", "filemenu1", language$, filemenu(1).Caption)
        ' filemenu(2) is a divider
        filemenu(3).Caption = gpps2("edit", "filemenu3", language$, filemenu(3).Caption)
            importmenu(0).Caption = gpps2("edit", "importmenu0", language$, importmenu(0).Caption)
            importmenu(1).Caption = gpps2("edit", "importmenu1", language$, importmenu(1).Caption)
        filemenu(4).Caption = gpps2("edit", "filemenu4", language$, filemenu(4).Caption)
            exportgaspermmenu.Caption = gpps2("edit", "exportgasperm", language$, exportgaspermmenu.Caption)
       '  filemenu(5) is a divider
        filemenu(6).Caption = gpps2("edit", "filemenu6", language$, filemenu(6).Caption)
    editmenu(2).Caption = gpps2("edit", "editmenu2", language$, editmenu(2).Caption)
    editmenu(3).Caption = gpps2("edit", "editmenu3", language$, editmenu(3).Caption)
    
    For i = 0 To 6
        set_fontstuff Pict_label(i), default_font
    Next i
    For i = 0 To 4
        set_fontstuff CItem(i), default_font
    Next i
    set_fontstuff gas_or_liquid, default_font
    
    For i = 0 To 16
        Label1(i).Caption = get_thing("edit", "label1" + str$(i), language$, Label1(i).Caption, Label1(i), default_font)
    Next i
    Option1(0).Caption = get_thing("edit", "option10", language$, Option1(0).Caption, Option1(0), default_font)
    Option1(1).Caption = get_thing("edit", "option11", language$, Option1(1).Caption, Option1(1), default_font)
    Check1.Caption = get_thing("edit", "check1", language$, Check1.Caption, Check1, default_font)
    
    ' Other text
    ts$(1) = gpps2("edit", "ts1", language$, "Default")
    ts$(2) = gpps2("edit", "ts2", language$, "Group Defined")
    ts$(3) = gpps2("edit", "ts3", language$, "Gas Permeameter")
    ts$(4) = gpps2("edit", "ts4", language$, "CFA")
    ts$(5) = gpps2("edit", "ts5", language$, "Integrity Tester")
    ts$(6) = gpps2("edit", "ts6", language$, "Bubble Point Tester")
    ts$(7) = gpps2("edit", "ts7", language$, "Capillary Flow")
    ts$(8) = gpps2("edit", "ts8", language$, "Data Editor")
    ts$(9) = gpps2("edit", "ts9", language$, "Unit")
    ts$(10) = gpps2("edit", "ts10", language$, "no file loaded")
    ts$(11) = gpps2("edit", "ts11", language$, "Gas Name")
    ts$(12) = gpps2("edit", "ts12", language$, "Fluid Name")
    ts$(13) = gpps2("edit", "ts13", language$, "Target Flow Rate (cc/min)")
    ts$(14) = gpps2("edit", "ts14", language$, "Chamber Volume (cc)")
    ts$(15) = gpps2("edit", "ts15", language$, "Do you want to close the editor without saving your changes?")
 '   ts$(16) = gpps2("edit", "ts16", language$, "Do you want to lose any changes you may have made to the current file?")
    ts$(17) = gpps2("edit", "ts17", language$, "Data File")
    ts$(18) = gpps2("edit", "ts18", language$, "Import")
    ts$(19) = gpps2("edit", "ts19", language$, "Error:  This data file does not contain a")
    ts$(20) = gpps2("edit", "ts20", language$, "input cancelled.")
    ts$(21) = gpps2("edit", "ts21", language$, "dry curve")
    ts$(22) = gpps2("edit", "ts22", language$, "wet curve")
    ts$(23) = gpps2("edit", "ts23", language$, "This edit operation is not possible.")
    ts$(24) = gpps2("edit", "ts24", language$, "Illegal Operation")
    'ts$(25) = gpps2("edit", "ts25", language$, "cc/min")
    ts$(26) = gpps2("edit", "ts26", language$, "Error: Permeability gas or liquid type not set properly")
    ts$(27) = gpps2("edit", "ts27", language$, "Error: Surface Tension incorrect")
    ts$(28) = gpps2("edit", "ts28", language$, "Error: Invalid diameter")
    ts$(29) = gpps2("edit", "ts29", language$, "Error: Invalid Cyl. Len.")
    ts$(30) = gpps2("edit", "ts30", language$, "Error: Invalid Thickness")
    ts$(31) = gpps2("edit", "ts31", language$, "Error: Invalid SC Diameter")
    ts$(32) = gpps2("edit", "ts32", language$, "Error: Invalid mass")
    ts$(33) = gpps2("edit", "ts33", language$, "Error: Invalid density")
    ts$(34) = gpps2("edit", "ts34", language$, "Error: Invalid efd flow")
    ts$(35) = gpps2("edit", "ts35", language$, "Error: Invalid chamber volume")
    ts$(36) = gpps2("edit", "ts36", language$, "Error: Invalid tortuosity factor")
    ts$(37) = gpps2("edit", "ts37", language$, "Please supply the sample diameter and thickness")
    ts$(38) = gpps2("edit", "ts38", language$, "(B)ubble Pt.")
    ts$(39) = gpps2("edit", "ts39", language$, "(T)rim")
    ts$(40) = gpps2("edit", "ts40", language$, "(Z)oom")
    ts$(41) = gpps2("edit", "ts41", language$, "(R)estore")
    ts$(42) = gpps2("edit", "ts42", language$, "(M)ove dry")
    ts$(43) = gpps2("edit", "ts43", language$, "Operator")
    ts$(44) = gpps2("edit", "ts44", language$, "Lot Number")
    ts$(45) = gpps2("edit", "ts45", language$, "Hardware Serial Number")
    ts$(46) = gpps2("edit", "ts46", language$, "Type of Test")
    ts$(47) = gpps2("edit", "ts47", language$, "Wet Parameters")
    ts$(48) = gpps2("edit", "ts48", language$, "Dry Parameters")
    ts$(49) = gpps2("edit", "ts49", language$, "Error: Invalid gas type")
    ts$(50) = gpps2("edit", "ts50", language$, "WARNING: You will lose any changes you may have made to the current file.")
    If hhRunAsBurst(current_unit%) Then
        ts$(51) = "Burst Tester Data File Editor"
        ts$(52) = "Burst Tester Data File"
    Else
        ts$(51) = gpps2("edit", "ts51", language$, "Hydrohead Tester Data File Editor")
        ts$(52) = gpps2("edit", "ts52", language$, "Hydrohead Tester Data File")
    End If
End Sub

