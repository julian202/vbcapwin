VERSION 5.00
Begin VB.Form frmAdditionalInformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Additional Information"
   ClientHeight    =   5010
   ClientLeft      =   11040
   ClientTop       =   5340
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5910
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   720
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   3600
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblNotice 
      Caption         =   $"frmAdditionalInformation.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "frmAdditionalInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private labels() As VB.Label
Private textboxes() As textbox

Private Sub SetupObjects()
    Dim i As Integer
    Dim numberOfInfoLines As Integer
    
    numberOfInfoLines = numberOfAdditionalInfoLines
    ReDim Preserve labels(numberOfInfoLines)
    ReDim Preserve textboxes(numberOfInfoLines)
    
    For i = 1 To numberOfInfoLines
        Set labels(i) = Me.Controls.Add("VB.Label", "Label" & i)
        Set textboxes(i) = Me.Controls.Add("VB.Textbox", "Textbox" & i)
        initLabel labels(i), i
        initTextbox textboxes(i), i
    Next i
    
    cmdCancel.top = labels(numberOfInfoLines).top + labels(numberOfInfoLines).Height + 240
    cmdSave.top = labels(numberOfInfoLines).top + labels(numberOfInfoLines).Height + 240
    Me.Height = cmdSave.top + cmdSave.Height + 670
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    Dim u$

    For i = 1 To numberOfAdditionalInfoLines
        infoLineValues(current_unit%, i - 1) = textboxes(i).Text
    Next i
    
    If current_unit% = 1 Then u$ = "" Else u$ = Format$(current_unit%)
    save_user_stuff u$
    
    Unload Me
End Sub

Private Sub Form_Load()
    SetupObjects
End Sub

Private Sub initLabel(ByRef Label As VB.Label, Index As Integer)
    With Label
        .Left = 240
        .top = 240 + (380 * (Index - 1))
        .Width = 2200
        .Height = 255
        .Visible = True
        .font.Size = 8
        .font.Bold = True
        .FontName = "MS Sans Serif"
        .Caption = infoLineHeaders(Index - 1)
    End With
End Sub

Private Sub initTextbox(ByRef textbox As VB.textbox, Index As Integer)
    With textbox
        .Left = 2240
        .top = 240 + (380 * (Index - 1))
        .Width = 3280
        .Height = 255
        .Visible = True
        .font.Size = 8
        .font.Bold = True
        .FontName = "MS Sans Serif"
        .Text = infoLineValues(current_unit%, Index - 1)
    End With
End Sub

