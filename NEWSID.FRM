VERSION 5.00
Begin VB.Form NewSID 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Group"
   ClientHeight    =   2175
   ClientLeft      =   2670
   ClientTop       =   3060
   ClientWidth     =   4335
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
   ScaleHeight     =   2175
   ScaleWidth      =   4335
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1665
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Commands 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   420
         Index           =   1
         Left            =   2475
         TabIndex        =   3
         Top             =   1155
         Width           =   1215
      End
      Begin VB.CommandButton Commands 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   420
         Index           =   0
         Left            =   675
         TabIndex        =   2
         Top             =   1155
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "New Group:"
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   1500
      End
   End
End
Attribute VB_Name = "NewSID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim ts$(9)

Private Sub Commands_Click(Index As Integer)

    Dim Ret$, i%, XTemp$, DFile$, numUsers%
    Dim Exists As Boolean, usernum%
    Dim LI As Integer
    Dim r As Long
    Dim fn As Integer
    Dim u$, unit%
    
    Ret$ = String$(255, vbNullChar)
    Exists = False
    If Index = 0 Then
        If Len(Text1.Text) <= 0 Then
            MsgBox ts$(1), 0, ts$(2)      '"The group name must be at least 1 character long."/"Group Name"
            Text1.SetFocus
            Exit Sub
        End If
        If Len(Text1.Text) > 255 Then
            MsgBox ts$(3), 0, ts$(2)          ' /"The group name must less than 256 characters long.""Group Name"
            Text1.SetFocus
            Exit Sub
        End If

        r = GPPS("default", "numusers", "1", Ret$, 255, IFile$)
        numUsers% = val(Ret$)

        For i% = 1 To numUsers%
            r = GPPS("default", "user" + Format$(i%), "", Ret$, 255, IFile$)
            If UCase$(nulltrim(Ret$)) = UCase$(Text1.Text) Then
                Exists = True
                usernum% = i%
                reply% = MsgBox(ts$(4) + vbCrLf + ts$(5), 36, ts$(6)) '"This name already exists."/"Replace?"/"In Use"
                If reply% = 6 Then
                    Exit For
                Else
                    Exit Sub
                End If
            End If
        Next i%
        If Not Exists Then
            ' add new user name
            XTemp$ = LTrim$(RTrim$(Text1.Text))
            On Error Resume Next
            MkDir EXE_Path$ + "users\" + XTemp$
            On Error GoTo bad_user_name
            fn = FreeFile
            Open EXE_Path$ + "users\" + XTemp$ + "\default.tpf" For Output As #fn
                Print #fn, str$(PULSEWIDTH): Print #fn, str$(V2INCR)
                Print #fn, str$(BUBLTIME): Print #fn, str$(BUBLFLOW)
                Print #fn, str$(ZEROTIME): Print #fn, str$(Maxpres)
                Print #fn, str$(MAXFLOW): Print #fn, str$(mineqtime)
                Print #fn, str$(PRESSLEW): Print #fn, str$(flowslew)
                Print #fn, str$(EQITER): Print #fn, str$(AVEITER)
                Print #fn, str$(MAXPDIF): Print #fn, str$(MAXFDIF)
                Print #fn, str$(STARTP): Print #fn, str$(STARTF)
                Print #fn, str$(PulseDelay): Print #fn, str$(preginc)
                Print #fn, str$(minbppres)
            Close #fn
            Rem copy the other files required from the current user
            If file_copy(XTemp$, "capgraph.ini") Or file_copy(XTemp$, "caprep.ini") Or file_copy(XTemp$, "rp_cover.txt") Then
                MsgBox ts$(7)               ' "Error copying files for new group"
                On Error GoTo 0
                Exit Sub
            End If
            On Error GoTo 0
            numUsers% = numUsers% + 1
            WPPS "default", "numusers", str$(numUsers%), IFile$
            WPPS "default", "user" + Format$(numUsers%), XTemp$, IFile$
            WPPS "default", "user", XTemp$, IFile$
            
            For unit% = 1 To chambers
                If unit% = 1 Then u$ = "" Else u$ = Format$(unit%)
                load_user_stuff u$ ' this loads globals when unit is 1
            Next unit%
            
            For unit% = 1 To chambers
                TPFWET$(unit%) = EXE_Path$ + "users\" + XTemp$ + "\default.tpf"
                TPFDRY$(unit%) = TPFWET$(unit%)
                check_group_change Curr_U$, XTemp$, OutFilename$(unit%)
            Next unit%
            Curr_U$ = XTemp$
            For unit% = 1 To chambers
                If unit% = 1 Then u$ = "" Else u$ = Format$(unit%)
                save_user_stuff u$ ' this saves globals when unit is 1
            Next unit%
        End If
    Else
    End If
done_user_name:
    On Error GoTo 0
    Unload Me
    Exit Sub
bad_user_name:
    MsgBox ts$(8)       ' "Error creating group name."
    Close #fn
    Resume done_user_name
    
End Sub

Private Sub Form_Load()

    LoadTextStrings
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(9) + str$(unitnumber)   ' "Unit"
    End If
    
    ' Superseded by LoadTextStrings
    'Text1.FontName = system_font
    'Text1.fontsize = font_size
    'Text1.fontbold = font_bold
    'edc 12-11-06 alter border color and caption
    'Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    Text1_GotFocus
    End Sub

Private Sub Text1_GotFocus()
'Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    ' new one only traps certain keys
    Select Case KeyAscii
        Case 0 To 7, 9 To 31, Asc("*"), Asc("?"), Asc("/"), Asc("\"), Asc("|"), Asc(":"), Asc("<"), Asc(">"), Asc("""")
            KeyAscii = 0
        Case 32
            If Len(Text1.Text) = 0 Then
                KeyAscii = 0
                ' can't have a space as first character
            End If
    End Select

End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    NewSID.Caption = gpps2("newSID", "window title", language$, NewSID.Caption)
    Label1.Caption = get_thing("newSID", "label1", language$, Label1.Caption, Label1, default_font)
    Commands(0).Caption = gpps2("newSID", "commands0", language$, Commands(0).Caption)
    set_fontname Commands(0), default_font
    Commands(1).Caption = gpps2("newSID", "commands1", language$, Commands(1).Caption)
    set_fontname Commands(1), default_font
    set_fontstuff Text1, default_font
    
    ' Other text
    ts$(1) = gpps2("newSID", "ts1", language$, "The group name must be at least 1 character long.")
    ts$(2) = gpps2("newSID", "ts2", language$, "Group Name")
    ts$(3) = gpps2("newSID", "ts3", language$, "The group name must less than 256 characters long.")
    ts$(4) = gpps2("newSID", "ts4", language$, "This name already exists.")
    ts$(5) = gpps2("newSID", "ts5", language$, "Replace?")
    ts$(6) = gpps2("newSID", "ts6", language$, "In Use")
    ts$(7) = gpps2("newSID", "ts7", language$, "Error copying files for new group")
    ts$(8) = gpps2("newSID", "ts8", language$, "Error creating group name.")
    ts$(9) = gpps2("newSID", "ts9", language$, "Unit")
    
End Sub
