VERSION 5.00
Begin VB.Form frmDebugWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug Screen"
   ClientHeight    =   3090
   ClientLeft      =   1680
   ClientTop       =   5025
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   12030
   Begin VB.ListBox lstDebugScreen 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
   End
End
Attribute VB_Name = "frmDebugWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddDebugStatement(statement As String)
'    If lstDebugScreen.ListCount > 1000 Then
'        lstDebugScreen.RemoveItem 0
'    End If
    
    lstDebugScreen.AddItem (Now & "     " & statement)
    lstDebugScreen.ListIndex = (lstDebugScreen.ListCount - 1)
End Sub
