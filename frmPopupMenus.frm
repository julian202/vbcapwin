VERSION 5.00
Begin VB.Form frmPopupMenus 
   Caption         =   "Popup Menus"
   ClientHeight    =   3090
   ClientLeft      =   7935
   ClientTop       =   9525
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8385
   Begin VB.Menu mnuPopupAdditionalInfo 
      Caption         =   "mnuPopupAdditionalInfo"
      Begin VB.Menu mnuPAIAdd 
         Caption         =   "Add..."
      End
      Begin VB.Menu mnuPAIInsert 
         Caption         =   "Insert..."
      End
      Begin VB.Menu mnuPAIRename 
         Caption         =   "Rename..."
      End
      Begin VB.Menu mnuPAIDelete 
         Caption         =   "Delete..."
      End
   End
End
Attribute VB_Name = "frmPopupMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuPAIAdd_Click()
    prefsForm.AddInfoLine
End Sub

Private Sub mnuPAIDelete_Click()
    prefsForm.DeleteInfoLine
End Sub

Private Sub mnuPAIInsert_Click()
    prefsForm.InsertInfoLine
End Sub

Private Sub mnuPAIRename_Click()
    prefsForm.RenameInfoLine
End Sub

