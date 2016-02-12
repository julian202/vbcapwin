VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Select_Test_New 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Selection"
   ClientHeight    =   6780
   ClientLeft      =   1890
   ClientTop       =   795
   ClientWidth     =   9405
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
   Icon            =   "SELTEST_NEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6780
   ScaleWidth      =   9405
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.Frame OptionsFrame 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Options"
         Height          =   4695
         Left            =   4440
         TabIndex        =   6
         Top             =   960
         Width           =   4455
      End
      Begin ComctlLib.TreeView OptionsTree 
         Height          =   4695
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   8281
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
      Begin VB.ComboBox cmbTestType 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "SELTEST_NEW.frx":030A
         Left            =   240
         List            =   "SELTEST_NEW.frx":032F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   3975
      End
      Begin VB.CommandButton ok 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Test Type"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Select_Test_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim genNode As Node
Dim testNode As Node
Dim subNode As Node

Private Sub cmbTestType_Click()
    Dim selItem As String
    
    selItem = cmbTestType.List(cmbTestType.ListIndex)
    
    OptionsTree.Nodes.clear
    addGenericOptions
    If selItem = "Capillary Flow Porometry" Then
        addCapFlowOptions
    ElseIf selItem = "Bubble Point" Then
        addBPOptions
    ElseIf selItem = "Gas Permeametry" Then
        addGPOptions
    ElseIf selItem = "Liquid Permeametry" Then
        addLPOptions
    ElseIf selItem = "Microflow" Then
        addMFOptions
    ElseIf selItem = "Gurley Test" Then
        addGUROptions
    ElseIf selItem = "Frazier Test" Then
        addFRAOptions
    End If
End Sub

Private Sub Form_Load()
    setupTestTypes
End Sub

Private Sub setupTestTypes()
    cmbTestType.clear
    If Not GasPerm And Not liqpermonly And Not itester And Not BPTester Then cmbTestType.AddItem "Capillary Flow Porometry"
    If Not GasPerm And Not liqpermonly And Not hydrohead_exclusive Then cmbTestType.AddItem "Bubble Point"
    If integrity And Not GasPerm And Not liqpermonly Then cmbTestType.AddItem "Integrity Test"
    If Not liqpermonly And Not itester And Not BPTester Then cmbTestType.AddItem "Gas Permeametry"
    If WESA_enabled And Not liqpermonly And Not itester And Not BPTester Then cmbTestType.AddItem "Envelope Surface Area"
    If Not liqpermonly And Not BPTester Then cmbTestType.AddItem "Pressure Hold"
    If DiffPG Then cmbTestType.AddItem "Microflow"
    If H2OPERM Then cmbTestType.AddItem "Liquid Permeametry"
    If hydrohead And Not burst Then cmbTestType.AddItem "Hydrohead"
    If hydrohead And burst Then cmbTestType.AddItem "Burst"
    If Not liqpermonly And Not itester And Not BPTester Then cmbTestType.AddItem "Gurley Test"
    If Not liqpermonly And Not itester And Not BPTester Then cmbTestType.AddItem "Frazier Test"
    If Resin_Diverter_Valve > 0 Then cmbTestType.AddItem "Resin Test"
    If hydrohead And mullen Then cmbTestType.AddItem "Mullen Test"
End Sub

Private Sub addGenericOptions()
    Set genNode = OptionsTree.Nodes.Add(, , "GenericRoot", "Generic Options")
    Set subNode = OptionsTree.Nodes.Add("GenericRoot", tvwChild, "grRunMultipleTests", "Run Multiple Tests")
    genNode.Expanded = True
End Sub

Private Sub addCapFlowOptions()
    Set testNode = OptionsTree.Nodes.Add(, , "CapFlowRoot", "Capillary Flow Options")
    Set subNode = OptionsTree.Nodes.Add("CapFlowRoot", tvwChild, "cfTestMethod", "Test Method")
    Set subNode = OptionsTree.Nodes.Add("CapFlowRoot", tvwChild, "cfMethodOptions", "Method Options")
    Set subNode = OptionsTree.Nodes.Add("CapFlowRoot", tvwChild, "cfAdditionalOptions", "Additional Options")
    testNode.Expanded = True
End Sub

Private Sub addBPOptions()
    Set testNode = OptionsTree.Nodes.Add(, , "BPRoot", "Bubble Point Options")
    Set subNode = OptionsTree.Nodes.Add("BPRoot", tvwChild, "bpDetectionMethod", "Detection Method")
    Set subNode = OptionsTree.Nodes.Add("BPRoot", tvwChild, "bpAdditionalOptions", "Additional Options")
    testNode.Expanded = True
End Sub

Private Sub addLPOptions()
    Set testNode = OptionsTree.Nodes.Add(, , "LPRoot", "Liquid Permeametry Options")
    Set subNode = OptionsTree.Nodes.Add("LPRoot", tvwChild, "lpTestMethod", "Test Method")
    testNode.Expanded = True
End Sub

Private Sub addGPOptions()
    Set testNode = OptionsTree.Nodes.Add(, , "GPRoot", "Gas Permeametry Options")
    Set subNode = OptionsTree.Nodes.Add("GPRoot", tvwChild, "gpTestMethod", "Test Method")
    Set subNode = OptionsTree.Nodes.Add("GPRoot", tvwChild, "gpAdditionalOptions", "Additional Options")
    testNode.Expanded = True
End Sub

Private Sub addMFOptions()
    Set testNode = OptionsTree.Nodes.Add(, , "MFRoot", "Microflow Options")
    Set subNode = OptionsTree.Nodes.Add("MFRoot", tvwChild, "mfTestOptions", "Test Options")
    testNode.Expanded = True
End Sub

Private Sub addGUROptions()
    Set testNode = OptionsTree.Nodes.Add(, , "GURRoot", "Gurley Options")
    Set subNode = OptionsTree.Nodes.Add("GURRoot", tvwChild, "gurTestMethod", "Test Method")
    testNode.Expanded = True
End Sub

Private Sub addFRAOptions()
    Set testNode = OptionsTree.Nodes.Add(, , "FRARoot", "Frazier Options")
    Set subNode = OptionsTree.Nodes.Add("FRARoot", tvwChild, "fraAdditionalOptions", "Additional Options")
    testNode.Expanded = True
End Sub

Private Sub OptionsTree_Click()
    Dim tmpTextBox As textbox
    Dim selNode As Node
    Dim nodeKey As String
    
    If OptionsTree.Nodes.count > 0 Then
        Set selNode = OptionsTree.SelectedItem
        nodeKey = selNode.Key
        If nodeKey = "grRunMultipleTests" Then
'            Set tmpTextBox = Controls.Add("vb.textbox", "tmpTextBox", OptionsFrame)
'
'            tmpTextBox.Width = 300
'            tmpTextBox.Height = 50
'            tmpTextBox.top = 500
'            tmpTextBox.Left = 500
'            tmpTextBox.Visible = True
        ElseIf nodeKey = "" Then
        
        End If
    End If
End Sub


'Generic Options
'    Run Multiple Tests
'        test count
'
'Capillary Flow Options
'    test Method
'        Wet Up/Dry Down
'            Stop Dry at Bubble Point
'        Wet Up/Dry Up
'            No Wait at Dry
'        Dry Up/Wet Up
'        wet Up / Calc.Dry
'            Linear, Darcy, or Square Root
'    Run Options
'        Run Normal
'        Run As microflow
'        Run as 3-Stage
'    Additional Options
'        Use Pressure Step List
'            Load Step List
'            Edit List
'        Run as Pass/Fail
'            Pass/Fail according to Bubble Point
'                Minimum pass bubble point pore size (microns)
'                Maximum pass bubble point pore size (microns)
'                Stop if test fails(acts as an abort command)
'            Pass/Fail according to Median Pore Size
'                Minimum pass Median pore size (microns)
'                Maximum pass Median pore size (microns)
'Bubble Point Options
'    Detection Method
'        manual
'        Auto Via F/PT
'        Auto Via Delta F/PT
'    Additional Options
'        Run as Pass/Fail
'            Minimum pass bubble point pore size (microns)
'            Maximum pass bubble point pore size (microns)
'            Stop if test fails
'
'Liquid Permeametry
'    Run Options
'        Ambient Pressure Test
'        Elevated Pressure Test
'        Build Cake By Pressure
'        Build Cake By Flow
'        Slurry Tube Wash Cycle
'
'Gas Permeametry
'    Run Options
'        Run Normal
'        Run As microflow
'        Run as 3-Stage
'    Additional Options
'        Use Pressure Step List
'            Load Step List
'            Edit List
'        Single Point Test
'
'Pressure Hold
'
'microflow
'    test Options
'        Linear Seal
'
'gurley
'    test Options
'        Standard Gurley Method
'        High Pressure Gurley
'        Sheffield Method
'
'frazier
'    Additional Options
'        Run as Pass/Fail
'
