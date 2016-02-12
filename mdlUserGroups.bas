Attribute VB_Name = "mdlUserGroups"
Option Explicit



Sub CheckUserGroups()
    Dim numIniUsers As Integer
    Dim numDirUsers As Integer
    Dim numMismatchIni As Integer
    Dim numMismatchDir As Integer
    Dim iniUsers() As String
    Dim dirUsers() As String
    Dim mismatchIni() As String
    Dim mismatchDir() As String
    Dim r As Long
    Dim Ret$
    Dim dirRet As String
    Dim i%
    Dim dirAttr As Long
    Dim errNum As Integer
    
    On Error GoTo UnknownErrorHandler
    errNum = 1

    Ret$ = String$(255, " ")
    r = GPPS("default", "numusers", "0", Ret$, 255, IFile$)
    numIniUsers = val(Ret$)
    ReDim iniUsers(numIniUsers)
    For i% = 1 To numIniUsers
        r = GPPS("default", "user" + Format$(i%), "", Ret$, 255, IFile$)
        iniUsers(i%) = UCase(nulltrim(Ret$))
    Next i%
    
    errNum = 2
    numDirUsers = 0
    dirRet = Trim$(Dir$(EXE_Path$ + "users\*", vbDirectory))
    While dirRet <> ""
        If dirRet <> "." And dirRet <> ".." Then
            dirAttr = GetAttr(EXE_Path$ + "users\" + dirRet)
            dirAttr = dirAttr And 16
            If dirAttr = 16 Then
                numDirUsers = numDirUsers + 1
                ReDim Preserve dirUsers(numDirUsers)
                dirUsers(numDirUsers) = UCase(nulltrim(dirRet))
            End If
        End If
        dirRet = Trim$(Dir$)
    Wend
    
    errNum = 3
    numMismatchIni = 0
    For i% = 1 To numIniUsers
        If Not existsInArray(dirUsers, iniUsers(i%)) Then
            numMismatchIni = numMismatchIni + 1
            ReDim Preserve mismatchIni(numMismatchIni)
            mismatchIni(numMismatchIni) = iniUsers(i%)
        End If
    Next i%
    
    errNum = 4
    numMismatchDir = 0
    For i% = 1 To numDirUsers
        If Not existsInArray(iniUsers, dirUsers(i%)) Then
            numMismatchDir = numMismatchDir + 1
            ReDim Preserve mismatchDir(numMismatchDir)
            mismatchDir(numMismatchDir) = dirUsers(i%)
        End If
    Next i%
    
    errNum = 5
    If numMismatchIni > 0 Or numMismatchDir > 0 Then
        MsgBox "There is a error in the group names between the initialization files and the user directories.  Please repair.", vbInformation
        Load frmGroupMismatch
        If numMismatchIni > 0 Then
            For i% = 1 To numMismatchIni
                frmGroupMismatch.lstINIEntries.AddItem (mismatchIni(i%))
            Next i%
            frmGroupMismatch.cmdIniDelete.Enabled = True
            frmGroupMismatch.cmdIniRename.Enabled = True
        End If
        If numMismatchDir > 0 Then
            For i% = 1 To numMismatchDir
                frmGroupMismatch.lstUserDirectories.AddItem (mismatchDir(i%))
            Next i%
            frmGroupMismatch.cmdDirDelete.Enabled = True
            frmGroupMismatch.cmdDirRename.Enabled = True
        End If
        frmGroupMismatch.Show vbModal
    End If
    errNum = 6
        
    Exit Sub
UnknownErrorHandler:
    MsgBox "Unknown Error " + vbCrLf + str$(Err.Number) + " " + Err.Description
End Sub

Function existsInArray(searchArray() As String, searchItem As String) As Boolean
    Dim uSize As Integer
    Dim i%
    
    On Error GoTo ErrorHandler
    existsInArray = False
    uSize = UBound(searchArray)
    For i% = 1 To uSize
        If searchArray(i%) = searchItem Then
            existsInArray = True
            Exit For
        End If
    Next i%
    Exit Function
ErrorHandler:
    MsgBox "Exists In Array Error" + vbCrLf + Err.Number + " " + Err.Description, vbInformation
End Function

