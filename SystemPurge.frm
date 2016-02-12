VERSION 5.00
Begin VB.Form SystemPurge 
   Caption         =   "System Purge"
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton buttContinue 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton buttNo 
      Caption         =   "NO"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton buttYes 
      Caption         =   "YES"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label labMessage 
      Alignment       =   2  'Center
      Caption         =   "Is the pison raised?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "SystemPurge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NextYesStep As String
Dim NextNoStep As String
Dim NextContStep As String
Dim ContinueClicked As Boolean

' Performs the next step based on which button is clicked.
' The step to perform is stored in the NextXXXXStep values.
Sub performStep(step As String)
    Dim command As String

    ' Get the step to perform based on the button clicked
    Select Case step
        Case "Y"
            command = NextYesStep
        Case "N"
            command = NextNoStep
        Case "C"
            command = NextContStep
        Case Else
            Exit Sub
    End Select
    
    ' Execute the command chosen
    Select Case command
    
        ' Raise the piston
        Case "RaisePiston"
            Move_Valve 14, "C"
            Set_Regulator 3, 1000
            
            ' Update UI
            labMessage.Caption = "Click CONTINUE once the piston has been raised."
            buttYes.Enabled = False
            buttNo.Enabled = False
            buttContinue.Enabled = True
            
            ' Setup button actions
            NextContStep = "StopPiston"
            
        ' Stop raising the piston
        Case "StopPiston"
            Set_Regulator 3, 0
            'Move_Valve 14, "O"
            GoTo RemoveSample
            
        ' Remove the sample
        Case "RemoveSample"
RemoveSample:
            ' Update UI
            labMessage.Caption = "Remove sample from liquid perm. chamber."
            buttYes.Enabled = False
            buttNo.Enabled = False
            buttContinue.Enabled = True
            
            ' Setup button actions
            NextContStep = "WaitPenetrometer"
            
        ' Wait for the penetrometer to reach a target
        Case "WaitPenetrometer"
            
            ' We need to track the next time the button is clicked
            ContinueClicked = False
            
            ' Update UI
            labMessage.Caption = "Purging." + Chr(13) + "Click CONTINUE to skip."
            buttYes.Enabled = False
            buttNo.Enabled = False
            buttContinue.Enabled = True
        
            ' Open valves
            Move_Valve 11, "O"
            Move_Valve 12, "O"
            
            ' Start timer
            StartTrackingTime
            Dim time As Long
            time = 0
            
            ' Read gauge
            ReadXReturnX4 6
            
            ' While the penetrometer is down, the purge has not been skipped, and time remains...
            Do While x4 > purgeStopCounts And Not ContinueClicked And time < 121
                labMessage.Caption = "Purging (" + CStr(120 - time) + " seconds remain, " + CStr(x4) + ")." + Chr(13) + "Click CONTINUE to skip."
                ReadXReturnX4 6
                time = GetSecondsSince
            Loop
            
            ' We no longer need to track this button
            ContinueClicked = True
            
            ' If we timed out
            If time > 120 Then
                ' Update the UI
                labMessage.Caption = "Add liquid to resevoir or pour directly into sample chamber."
                buttYes.Enabled = False
                buttNo.Enabled = False
                buttContinue.Enabled = True
                
                ' Setup button actions
                NextContStep = "WaitPenetrometer"
                
                Exit Sub
            End If
            
            ' Close the valves
            Move_Valve 11, "C"
            Move_Valve 12, "C"
            
            ' Update the UI
            labMessage.Caption = "Is the liquid filling the sample chamber?"
            buttYes.Enabled = True
            buttNo.Enabled = True
            buttContinue.Enabled = False
            
            ' Setup button actions
            NextYesStep = "Exit"
            NextNoStep = "PourLiquid"
            
        Case "PourLiquid"
            ' Update the UI
            labMessage.Caption = "Pour liquid to fill sample chamber."
            buttYes.Enabled = False
            buttNo.Enabled = False
            buttContinue.Enabled = True
            
            ' Setup button actions
            NextContStep = "Exit"
            
        Case "Exit"
            Unload Me
            
    End Select
    
End Sub

'Sub Set_Regulator(channel As Integer, counts As Integer)
'    Send_RS232 ("Z" + CStr(channel))
'
'    Do While counts > 255
'        Send_RS232b "U" + CStr(channel), 255
'        counts = counts - 255
'    Loop
'
'    If counts > 0 Then
'        Send_RS232b "U" + CStr(channel), CByte(counts)
'    End If
'End Sub

Sub StartTrackingTime()
    TrackTime = Timer
End Sub

Function GetSecondsSince() As Long
    Dim now As Long
    
    now = Timer
    
    If now < TrackTime Then TrackTime = (60 * 60 * 12) - TrackTime
    
    GetSecondsSince = now - TrackTime
End Function

Private Sub buttContinue_Click()
    ' Only perform the step if we're not listening to that button for other reasons
    If ContinueClicked Then
        performStep ("C")
    Else
        ContinueClicked = True
        buttContinue.Enabled = False
    End If
End Sub

Private Sub buttNo_Click()
    performStep ("N")
End Sub

Private Sub buttYes_Click()
    performStep ("Y")
End Sub

Private Sub Form_Load()
    NextYesStep = "RemoveSample"
    NextNoStep = "RaisePiston"
    ContinueClicked = True
End Sub
