VERSION 5.00
Begin VB.Form stepListForm 
   BackColor       =   &H000000FF&
   Caption         =   "Pressure Step List Editor"
   ClientHeight    =   7740
   ClientLeft      =   2820
   ClientTop       =   1245
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.ListBox stepList 
         Height          =   4740
         ItemData        =   "stepListForm.frx":0000
         Left            =   120
         List            =   "stepListForm.frx":0002
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   4080
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   5520
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear List"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label1 
         Height          =   3375
         Left            =   3000
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Menu fileMenu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu newMenu 
         Caption         =   "&New..."
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu openMenu 
         Caption         =   "&Open..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu blank 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu saveMenu 
         Caption         =   "&Save"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu saveAsMenu 
         Caption         =   "Save &As..."
         Index           =   1
      End
      Begin VB.Menu blank2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu exitMenu 
         Caption         =   "E&xit"
         Index           =   1
      End
   End
End
Attribute VB_Name = "stepListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim currentList As pressureList
Dim listFilePath$
Dim ts$(8)                              ' Text forms for this module

Private Sub Command1_Click()
' User has clicked the "add" button

    Dim textToAdd$
    Dim i As Integer
    Dim results As SplitStringResult

    ' First determine whether we have a single value or several
    textToAdd$ = Text1.Text
    If textToAdd$ = "" Then Exit Sub
    i = InStr(textToAdd$, ";")
    
    If i = 0 Then               ' Single value on the line
        checkAndAdd textToAdd$
    Else                        ' Multiple values
        results = splitString(textToAdd$, ";")
        For i = 1 To results.numberOfValues
            checkAndAdd results.values(i)
        Next i
    End If
    
    Text1.Text = ""
    currentList.dirty = True
    
    
End Sub

Private Sub checkAndAdd(a$)
' Check to see whether the item already exists in the list. If so, do nothing; else add it in order

    Dim i As Integer
    Dim lessThanIndex As Integer    ' Index of greatest value in list less than value being added
    
    lessThanIndex = -1
    For i = 0 To stepList.ListCount - 1
        If stepList.List(i) = a$ Then Exit Sub
        If val(stepList.List(i)) < val(a$) Then lessThanIndex = i
    Next i
    
    stepList.AddItem a$, lessThanIndex + 1
    currentList.dirty = True

End Sub

Private Sub Command2_Click()
' Delete the pressure(s) selected in the list

    Dim i As Integer

    ' We go backwards here because if we delete as we go, in order, we'll change the
    ' number of items and get an index error at the end (at the very least). Reverse order,
    ' however, works just fine.
    For i = (stepList.ListCount - 1) To 0 Step -1
        If stepList.Selected(i) Then stepList.RemoveItem (i)
    Next i
    
    currentList.dirty = True
        
End Sub

Private Sub Command3_Click()
' Clear list

    stepList.clear
    currentList.dirty = True
    
End Sub

Private Sub exitMenu_Click(Index As Integer)
' Get outta here

    ' Check for unsaved list
    checkDirtySave
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    LoadTextStrings
    Label2.Caption = ts$(1) + " (" + PU$ + ")"   ' "Pressure"
    Label1.Caption = ts$(7) + " " + ts$(8)
    'edc 12-11-06 alter border color and caption
    Me.Caption = Me.Caption & "    " & SubCaption
    Me.BackColor = lngBorderColor
    
End Sub

Private Sub newMenu_Click(Index As Integer)
' Creating a new file. If the old one is dirty, prompt for a save - else
' clear the array, clear the window.

    checkDirtySave          ' Save if dirty and user wishes
    
    With currentList
        .dirty = False
        .count = 0
        ReDim .values(0)
    End With
    
    stepList.clear
    Text1.Text = ""
    
End Sub

Private Sub openMenu_Click(Index As Integer)
' Let user select a list to open

    Dim result As Boolean
    Dim i As Integer
    
    ' Configure and open file selection box
    fsel_path$ = EXE_Path$ + "\parms\*.psl"
    fsel_title$ = ts$(2)            ' "Choose pressure list"
    fsel_name$ = ""
    fsel_io = True
    fsel Me.hwnd
   
    If fsel_return = "" Then Exit Sub
    
    listFilePath$ = fsel_return
    
    ' Load the file into currentList
    result = loadList(listFilePath$, currentList)
    
    If Not result Then Exit Sub
    
    For i = 1 To currentList.count
        checkAndAdd Format$(currentList.values(i))
    Next i
    
End Sub

Private Function loadList(filename$, aList As pressureList) As Boolean
' Load the values from a pressure list in filename$ into the pressure list specified

    Dim i, fn1 As Integer
    Dim temp$
    
    fn1 = FreeFile
    Open filename$ For Input As #fn1
    
    Input #fn1, temp$
    If Left$(temp$, 1) <> "#" Then
        MsgBox (ts$(3))     ' "Error: The selected file is not a valid pressure step list."
        loadList = False
        Exit Function
    End If
    
    aList.count = val(Right(temp$, Len(temp$) - 1))     ' Fixed 6.71.61 - was only taking last character of number
    ReDim aList.values(aList.count)
    
    On Error GoTo errortrap
    
    For i = 1 To aList.count
        Input #fn1, aList.values(i)
    Next i
    
    Close
    aList.dirty = False
    
    loadList = True
    Exit Function

errortrap:
    Close
    MsgBox (ts$(4))         ' "Error: The selected pressure list is corrupted."
    loadList = False

End Function

Private Sub saveAsMenu_Click(Index As Integer)
' Let the user pick a place to save the file, then call saveList to do the grunt work.

    fsel_name$ = ""
    fsel_title$ = ts$(5)        ' "Save pressure step list"
    fsel_path$ = EXE_Path$ + "parms\*.psl"
    fsel_io = False
    fsel Me.hwnd
    If fsel_return$ <> "" Then                  ' valid filename
        listFilePath$ = fsel_return$            ' Selected path becomes the default
        saveList listFilePath$
    End If
                
End Sub

Private Sub saveMenu_Click(Index As Integer)
' Save to the current pathname, if there is one
    
    If listFilePath <> "" Then
        saveList listFilePath$
    Else
        saveAsMenu_Click (1)                    ' Call up the "save as" dialog
    End If
    
End Sub

Private Sub saveList(path$)
' Save currentList to path$

    Dim i, fn1 As Integer
    Dim temp$

    If path$ <> "" Then
        ' Saving consists of rewriting the file from scratch
        fn1 = FreeFile
        Open path$ For Output As #fn1
        
        temp$ = "#" + Format$(stepList.ListCount)       ' This is the number of steps in the file
        Print #fn1, temp$
        For i = 0 To stepList.ListCount - 1
            Print #fn1, stepList.List(i)
        Next i
        
        Close
        
        currentList.dirty = False
        
    End If
    
End Sub

Private Sub checkDirtySave()
' If currentList is dirty, prompt the user to save the file if desired. Then do the save.

    Dim result As Integer
    
    If Not currentList.dirty Then Exit Sub
    
    result = MsgBox(ts$(6), vbYesNo)        ' "The current pressure list has been modified. Do you wish to save the list before continuing?"
    
    If result = vbYes Then
        saveMenu_Click (1)          ' Call the standard save routine
    End If

End Sub

Public Sub LoadTextStrings()
' Load text elements for this form from external .ini file
    
    Dim i As Integer

    ' Form elements
    stepListForm.Caption = get_thing("steplist", "window title", language$, stepListForm.Caption, stepListForm, default_font)
    Label2.Caption = get_thing("steplist", "label2", language$, Label2.Caption, Label2, default_font)
    set_fontstuff Label1, default_font          ' set in code
    set_fontstuff Command1, default_font
    set_fontstuff Command2, default_font
    set_fontstuff Command3, default_font
    Command1.Caption = gpps2("steplist", "command1", language$, Command1.Caption)
    Command2.Caption = gpps2("steplist", "command2", language$, Command2.Caption)
    Command3.Caption = gpps2("steplist", "command3", language$, Command3.Caption)
    set_fontstuff stepList, default_font
    set_fontname Text1, default_font
    filemenu(0).Caption = gpps2("steplist", "filemenu", language$, filemenu(0).Caption)
    exitmenu(1).Caption = gpps2("steplist", "exitmenu", language$, exitmenu(1).Caption)
    newMenu(1).Caption = gpps2("steplist", "newmenu", language$, newMenu(1).Caption)
    openMenu(1).Caption = gpps2("steplist", "openmenu", language$, openMenu(1).Caption)
    saveMenu(1).Caption = gpps2("steplist", "savemenu", language$, saveMenu(1).Caption)
    saveAsMenu(1).Caption = gpps2("steplist", "saveasmenu", language$, saveAsMenu(1).Caption)
    
    
    ' Other text
    ts$(1) = gpps2("steplist", "ts1", language$, "Pressure")
    ts$(2) = gpps2("steplist", "ts2", language$, "Choose pressure list")
    ts$(3) = gpps2("steplist", "ts3", language$, "Error: The selected file is not a valid pressure step list.")
    ts$(4) = gpps2("steplist", "ts4", language$, "Error: The selected pressure list is corrupted.")
    ts$(5) = gpps2("steplist", "ts5", language$, "Save pressure step list")
    ts$(6) = gpps2("steplist", "ts6", language$, "The current pressure list has been modified. Do you want to save the list before continuing?")
    ts$(7) = gpps2("steplist", "ts7", language$, "Data will be acquired only at the pressures specified in the list.  To add values, enter single pressures or a semicolon-delimited list and click 'Add'.")
    ts$(8) = gpps2("steplist", "ts8", language$, "To remove values, highlight one or more pressures in the list and click 'Delete'.  Click on 'Clear list' to remove all pressures from the list.")
    
End Sub
