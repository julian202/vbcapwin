VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form BorderColor 
   BackColor       =   &H000000FF&
   Caption         =   "Change Border Color"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdOK 
         Caption         =   "Accept and Apply"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   2640
         Width           =   2415
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton change 
         Caption         =   "Change Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cancel 
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
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "The current border color is : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "To change this setting Click the ""Change Color"" button below. To exit this dialog click the ""Cancel"" button."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   3975
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "BorderColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancel_Click()
    Me.Hide
End Sub

Private Sub change_Click()
    Dim i As Long

    CommonDialog1.CancelError = True
On Error GoTo errhandler
    CommonDialog1.flags = cdlCCRGBInit
    CommonDialog1.ShowColor
    Me.BackColor = CommonDialog1.Color
    lngBorderColor = Me.BackColor
    'i = WPPS("default", "border", lngBorderColor, Ifile$)
    Picture1.BackColor = lngBorderColor
    Exit Sub

errhandler:
    Select Case Err
    Case 32755 '  Dialog Cancelled
        MsgBox "you cancelled the dialog box"
    Case Else
        MsgBox "Unexpected error. Err " & Err & " : " & error
    End Select
End Sub

Private Sub cmdOK_Click()
    lngBorderColor = Me.BackColor
    TitleScrn.BackColor = lngBorderColor
    WPPS "Capstuff", "border", lngBorderColor, CSFile$
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.BackColor = lngBorderColor
    Picture1.BackColor = lngBorderColor
End Sub
