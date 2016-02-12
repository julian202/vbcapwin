VERSION 5.00
Begin VB.Form frmGroupMismatch 
   BackColor       =   &H000000FF&
   Caption         =   "Group Mismatch"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton cmdDirDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   8
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdDirRename 
         BackColor       =   &H00FFFF80&
         Caption         =   "Rename"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdIniDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdIniRename 
         BackColor       =   &H00FFFF80&
         Caption         =   "Rename"
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ListBox lstUserDirectories 
         Height          =   2985
         Left            =   4800
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
      Begin VB.ListBox lstINIEntries 
         Height          =   2985
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF80&
         Caption         =   "User Directories:"
         Height          =   255
         Left            =   4800
         TabIndex        =   4
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF80&
         Caption         =   "INI Entries:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmGroupMismatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDirDelete_Click()
    Dim msgres As VbMsgBoxResult
    Dim userDir As String
    
    On Error GoTo ErrorHandler
    
    If lstUserDirectories.ListIndex >= 0 Then
        userDir = lstUserDirectories.List(lstUserDirectories.ListIndex)
        msgres = MsgBox("Are you sure you want to delete the " + userDir + " directory?", vbYesNo)
        If msgres = vbYes Then
            On Error Resume Next
            Kill EXE_Path$ + "users\" + userDir + "\*.*"
            On Error GoTo 0
            RmDir EXE_Path$ + "users\" + userDir
            If Dir(EXE_Path$ + "users\" + userDir) <> "" Then
                MsgBox "Unable to delete " + userDir + " directory", vbInformation
            Else
                lstUserDirectories.RemoveItem lstUserDirectories.ListIndex
            End If
        End If
    Else
        MsgBox "Please select a User Directory first.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "cmdDirDelete_Click Error" + vbCrLf + Err.Number + " " + Err.Description
End Sub

Private Sub cmdDirRename_Click()
    Dim newName As String
    Dim userDir As String
    
    On Error GoTo ErrorHandler
    
    If lstUserDirectories.ListIndex >= 0 Then
        userDir = lstINIEntries.List(lstINIEntries.ListIndex)
        If Dir(EXE_Path$ + "users\" + userDir, vbDirectory) = "" Then
            MsgBox "Directory " + userDir + " does not exist.  Unable to rename.", vbInformation
        Else
            newName = UCase$(InputBox("Please enter a new directory name", "Rename Directory", userDir))
            If newName <> "" Then
                If Dir(EXE_Path$ + "users\" + newName, vbDirectory) <> "" Then
                    MsgBox "Directory " + newName + " already exists. Unable to rename.", vbInformation
                Else
                    Name EXE_Path$ + "users\" + userDir As EXE_Path$ + "users\" + newName
                    If Dir(EXE_Path$ + "users\" + newName, vbDirectory) = "" Then
                        MsgBox "Unable to rename " + userDir + " to " + newName + ".", vbInformation
                    Else
                        lstUserDirectories.List(lstUserDirectories.ListIndex) = newName
                    End If
                End If
            End If
        End If
    Else
        MsgBox "Please select a User Directory first.", vbInformation
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "cmdDirRename_Click Error" + vbCrLf + Err.Number + " " + Err.Description
End Sub

Private Sub cmdIniDelete_Click()
    Dim r As Long
    Dim Ret$
    Dim numUsers%
    Dim usernum%
    Dim userIni As String
    Dim i%
    
    On Error GoTo ErrorHandler
    
    If lstINIEntries.ListIndex >= 0 Then
        userIni = lstINIEntries.List(lstINIEntries.ListIndex)
        Ret$ = String$(255, vbNullChar)
        r = GPPS("default", "numusers", "1", Ret$, 255, IFile$)
        numUsers% = val(Ret$)
        usernum% = 0
        For i% = 1 To numUsers%
            r = GPPS("default", "user" + Format$(i%), "", Ret$, 255, IFile$)
            If UCase$(nulltrim(Ret$)) = UCase$(userIni) Then
                usernum% = i%
                Exit For
            End If
        Next i%
        If usernum% = 0 Then
            MsgBox "Unable to find group name"                       ' "Error:  Couldn't find group name in index"
        Else
            If usernum% < numUsers% Then
                ' swap last user into place of deleted one
                r = GPPS("default", "user" + Format$(numUsers%), "", Ret$, 255, IFile$)
                WPPS "default", "user" + Format$(usernum%), nulltrim(Ret$), IFile$
                usernum% = numUsers%
            End If
            WPPS userIni, vbNullString, vbNullString, IFile$
            WPPS "default", "user" + Format$(usernum%), vbNullString, IFile$
            numUsers% = numUsers% - 1
            WPPS "default", "numusers", str$(numUsers%), IFile$
            lstINIEntries.RemoveItem lstINIEntries.ListIndex
        End If
    Else
        MsgBox "Please select an INI entry first.", vbInformation
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "cmdIniDelete_Click Error" + vbCrLf + Err.Number + " " + Err.Description
End Sub

Private Sub cmdIniRename_Click()
    Dim newName As String
    Dim Ret$
    Dim numUsers%
    Dim usernum%
    Dim unit%
    Dim u$
    Dim XTemp$
    Dim userIni As String
    Dim Exists As Boolean
    Dim r As Long
    Dim i%
    
    On Error GoTo ErrorHandler
    
    If lstINIEntries.ListIndex >= 0 Then
        userIni = lstINIEntries.List(lstINIEntries.ListIndex)
        newName = Trim(UCase$(InputBox("Please enter a new name for the INI entry", "Rename INI", userIni)))
        If newName <> "" Then
            Exists = False
            usernum% = 0
            Ret$ = String$(255, vbNullChar)
            r = GPPS("default", "numusers", "1", Ret$, 255, IFile$)
            numUsers% = val(Ret$)
            For i% = 1 To numUsers%
                r = GPPS("default", "user" + Format$(i%), "", Ret$, 255, IFile$)
                If UCase$(nulltrim(Ret$)) = UCase$(userIni) Then
                    usernum% = i%
                End If
                If UCase$(nulltrim(Ret$)) = newName Then
                    MsgBox "New name already exists.  Unable to rename."          ' "New name already exists - name change aborted"
                    Exists = True
                    Exit For
                End If
            Next i%
            If usernum% = 0 Then
                MsgBox "Old name not found.  Unable to rename."           ' "Something is wrong - current user index couldn't be found"
                Exists = True ' set to true to bypass changing name
            End If
            If Not Exists Then
                XTemp$ = newName
                For unit% = 1 To chambers
                    If unit% = 1 Then u$ = "" Else u$ = Format$(unit%)
                    load_user_stuff u$
                Next unit%
                WPPS userIni, vbNullString, vbNullString, IFile$
                WPPS "default", "user", XTemp$, IFile$
                WPPS "default", "user" + Format$(usernum%), XTemp$, IFile$
                ' when renaming a user, if the parameters point to the old user directory,
                ' they now need to point to the new user directory
                For unit% = 1 To chambers
                    check_group_change userIni, XTemp$, TPFDRY$(unit%)
                    check_group_change userIni, XTemp$, TPFWET$(unit%)
                    check_group_change userIni, XTemp$, OutFilename$(unit%)
                Next unit%
                check_group_change userIni, XTemp$, path(0)
                check_group_change userIni, XTemp$, path(1)
                
                For unit% = 1 To chambers
                    If unit% = 1 Then u$ = "" Else u$ = Format$(unit%)
                    save_user_stuff u$
                Next unit%
                
                lstINIEntries.List(lstINIEntries.ListIndex) = XTemp$
            End If
        End If
    Else
        MsgBox "Please select an INI entry first.", vbInformation
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "cmdIniRename_Click Error" + vbCrLf + Err.Number + " " + Err.Description
End Sub
