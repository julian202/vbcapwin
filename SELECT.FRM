VERSION 5.00
Begin VB.Form Selection 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECT"
   ClientHeight    =   3510
   ClientLeft      =   1530
   ClientTop       =   2475
   ClientWidth     =   8175
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
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3510
   ScaleWidth      =   8175
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.ListBox Obj_Select 
         BackColor       =   &H00FFFFFF&
         Columns         =   3
         Height          =   2205
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   7695
      End
      Begin VB.CommandButton SelCancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton SelCont 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   2520
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefSng A-Z
Dim ts$(7)              ' Text strings for this form
Dim unsort%()

Private Sub Form_Load()

    Dim Ret$, i%, numUsers%, startuser%, a$
    Dim r As Long, maxwidth As Integer, j%, k%
    
    LoadTextStrings

    If T_Select$ = "DEL" Then SelCont.Caption = ts$(1)      ' "&Delete"
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.top = (Screen.Height - Me.Height) / 2
    Obj_Select.clear
    Ret$ = String$(255, vbNullChar)
    If T_Select$ = "USER" Or T_Select$ = "DEL" Then
        r = GPPS("default", "numusers", "1", Ret$, 255, IFile$)
        numUsers% = val(Ret$)
        maxwidth = 0
        j% = -1 ' current user not found yet
        If T_Select$ = "DEL" Then
            startuser% = 2
            If (numUsers% = 1) Or (numUsers% = 2 And UCase$(Curr_U$) <> "DEFAULT") Then
                a$ = ts$(2)                             ' "There are no groups that you can delete."
                maxwidth = Me.TextWidth(a$)
                Obj_Select.AddItem a$ + vbTab + Format$(0)
                Obj_Select.Enabled = False
                SelCont.Enabled = False
            End If
        Else
            startuser% = 1
        End If
        k% = 0 ' offset for skipped entry - haven't skipped it yet
        For i% = startuser% To numUsers%
            r = GPPS("default", "user" + Format$(i%), "", Ret$, 255, IFile$)
            a$ = UCase$(nulltrim(Ret$))
            If UCase$(Curr_U$) = UCase$(a$) Then
                j% = i% - startuser%
            End If
            If startuser% <> 2 Or j% <> i% - startuser% Then
                ' don't put it in if this is the current user and we are deleting
                r = Me.TextWidth(a$)
                If maxwidth < r Then maxwidth = r
                ' add the original index number to the end
                ' so we can recover the sorted order
                Obj_Select.AddItem a$ + vbTab + Format$(i% - startuser% - k%)
            ElseIf startuser% = 2 Then
                k% = 1
            End If
        Next i%
        
        ReDim unsort%(Obj_Select.ListCount - 1)
        For i% = 0 To Obj_Select.ListCount - 1
            a$ = Obj_Select.List(i%)
            k% = InStr(1, a$, vbTab)
            Obj_Select.ItemData(i%) = val(Mid$(a$, k% + 1))
            Obj_Select.List(i%) = Left$(a$, k% - 1)
            unsort%(Obj_Select.ItemData(i%)) = i%
        Next i%
        Obj_Select.Columns = Obj_Select.Width \ maxwidth
        'If startuser% <> 2 Then Obj_Select.ListIndex = unsort%(j%)
        If T_Select$ = "GAS" Then
            numUsers% = 1 ' this is a local variable, so I can reuse it
            startuser% = -1 ' haven't found current gas yet
            maxwidth = 0
            Do
                r = GPPS("main", "gas" + Format$(numUsers%), "", Ret$, 255, EXE_Path$ + "gasflowconversion.ini")
                If r = 0 Then Exit Do
                a$ = nulltrim(Ret$)
                r = Me.TextWidth(a$)
                If maxwidth < r Then maxwidth = r
                Obj_Select.AddItem a$ + vbTab + Format$(numUsers% - 1)
                If a$ = TitleScrn.GasLabel.Caption Then
                    startuser% = numUsers% - 1
                End If
                numUsers% = numUsers + 1
            Loop
            numUsers% = numUsers - 1
            If numUsers% < 2 Then
                Obj_Select.Enabled = False
                SelCont.Enabled = False
                ' they can't select anything if there is only one
            ElseIf startuser% >= 0 Then
            ReDim unsort%(numUsers% - 1)
                For i% = 0 To numUsers% - 1
                    a$ = Obj_Select.List(i%)
                    k% = InStr(1, a$, vbTab)
                    Obj_Select.ItemData(i%) = val(Mid$(a$, k% + 1))
                    Obj_Select.List(i%) = Left$(a$, k% - 1)
                    unsort%(Obj_Select.ItemData(i%)) = i%
                Next i%
                Obj_Select.ListIndex = unsort%(startuser%)
                Obj_Select.Columns = Obj_Select.Width \ maxwidth
            End If
        End If
    End If
    If unitnumber <> 0 Then
        Me.Caption = Me.Caption + " - " + ts$(3) + str$(unitnumber)       '"Unit"
    End If
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    Me.Obj_Select.clear
End Sub

Private Sub Obj_Select_DblClick()
    SelCont_Click
End Sub

Private Sub SelCancel_Click()
    T_Select$ = "Cancel"
    Unload Me
End Sub

Private Sub SelCont_Click()

    Dim To_Del_U$, temp$, i%, Ret$, numUsers%, usernum%
    Dim r As Long
    
    Ret$ = String$(255, vbNullChar)
    If T_Select$ = "DEL" Then
        If UCase$(Curr_U$) = UCase$(Obj_Select.List(Obj_Select.ListIndex)) Then
            MsgBox ts$(4), 0, ts$(6)               ' "The current Group cannot be deleted."/"Group"
        Else
            To_Del_U$ = UCase$(Obj_Select.List(Obj_Select.ListIndex))
            Obj_Select.RemoveItem Obj_Select.ListIndex
            r = GPPS("default", "numusers", "1", Ret$, 255, IFile$)
            numUsers% = val(Ret$)
            usernum% = 0
            For i% = 1 To numUsers%
                r = GPPS("default", "user" + Format$(i%), "", Ret$, 255, IFile$)
                If UCase$(nulltrim(Ret$)) = UCase$(To_Del_U$) Then
                    usernum% = i%
                    Exit For
                End If
            Next i%
            If usernum% = 0 Then
                MsgBox ts$(5)                       ' "Error:  Couldn't find group name in index"
            Else
                If usernum% < numUsers% Then
                    ' swap last user into place of deleted one
                    r = GPPS("default", "user" + Format$(numUsers%), "", Ret$, 255, IFile$)
                    WPPS "default", "user" + Format$(usernum%), nulltrim(Ret$), IFile$
                    usernum% = numUsers%
                End If
                WPPS To_Del_U$, vbNullString, vbNullString, IFile$
                WPPS "default", "user" + Format$(usernum%), vbNullString, IFile$
                numUsers% = numUsers% - 1
                WPPS "default", "numusers", str$(numUsers%), IFile$
                ' we could delete the files if necessary
            End If
        End If
        Unload Me
        Exit Sub
    End If
    If T_Select$ = "USER" Then
        Curr_U$ = UCase$(Obj_Select.List(Obj_Select.ListIndex))
    End If
    If T_Select$ = "GAS" Then
        If Obj_Select.ListIndex < 0 Then Exit Sub ' don't allow null selection
        i% = Obj_Select.ItemData(Obj_Select.ListIndex)
        TitleScrn.GasLabel.Caption = Obj_Select.List(Obj_Select.ListIndex)
        GPPS "main", "conversion" + Format$(i% + 1), "1", Ret$, 255, EXE_Path$ + "gasflowconversion.ini"
        gasflowconversionfactor = val(Ret$)
        If gasflowconversionfactor <= 0 Then gasflowconversionfactor = 1
        WPPS "main", "current", Format$(i% + 1), EXE_Path$ + "gasflowconversion.ini"
    End If
    If LTrim$(RTrim$(Curr_U$)) = "" Then
        MsgBox ts$(7), 0, ts$(6)         ' "Group"/' "Select a Group"
        Exit Sub
    End If
    Unload Me

End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    ' Form elements
    Selection.Caption = gpps2("selection", "window title", language$, Selection.Caption)
    set_fontstuff Obj_Select, default_font
    SelCont.Caption = gpps2("selection", "continue", language$, SelCont.Caption)
    set_fontname SelCont, default_font
    SelCancel.Caption = gpps2("selection", "cancel", language$, SelCancel.Caption)
    set_fontname SelCancel, default_font
    
    ' Other text
    ts$(1) = gpps2("selection", "ts1", language$, "&Delete")
    ts$(2) = gpps2("selection", "ts2", language$, "There are no groups that you can delete.")
    ts$(3) = gpps2("selection", "ts3", language$, "Unit")
    ts$(4) = gpps2("selection", "ts4", language$, "The current group cannot be deleted.")
    ts$(5) = gpps2("selection", "ts5", language$, "Error:  Couldn't find group name in index")
    ts$(6) = gpps2("selection", "ts6", language$, "Group")
    ts$(7) = gpps2("selection", "ts7", language$, "Select a Group")
    
End Sub

