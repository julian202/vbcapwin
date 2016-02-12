VERSION 5.00
Begin VB.Form GLSel 
   BackColor       =   &H000000FF&
   Caption         =   "Gas/Liquid Select"
   ClientHeight    =   4695
   ClientLeft      =   1905
   ClientTop       =   2355
   ClientWidth     =   7230
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
   LinkTopic       =   "GLSel"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4695
   ScaleWidth      =   7230
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtConvFactor 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2640
         TabIndex        =   16
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtNome 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   " "
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txtVisc 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2640
         TabIndex        =   5
         Text            =   " "
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox cboGas 
         Height          =   315
         ItemData        =   "GLSEL.frx":0000
         Left            =   1560
         List            =   "GLSEL.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox cboLiqid 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton cmdGas 
         Caption         =   "Add a Gas"
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         ToolTipText     =   "To Add a gas click on this button fill in the name and viscoity then click the Commit Changes to finish process"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdLiqid 
         Caption         =   "Add a Liquid"
         Height          =   375
         Left            =   4560
         TabIndex        =   3
         ToolTipText     =   $"GLSEL.frx":0004
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdComit 
         Caption         =   "Commit Changes"
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txtSur 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   2640
         TabIndex        =   6
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Gas Flow Conversion Factor:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Name of Gas or Liquid:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Viscosity (cP):"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Select a Gas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Select a Liquid"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Surface Tension dynes/cm:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   2415
      End
   End
End
Attribute VB_Name = "GLSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ts$(3) ' Text strings for this form
Dim dump() As String
Dim numGL As Integer
Dim gasNum, liqNum, fluNum, indexF, indexG, indexL, lenFl, lenGas, lenLiq As Integer ' Declaring various variables
Dim targetIsLiqid As Boolean
Dim addGL, clear As Boolean
Dim GasAray() As String     'An array for the gas
Dim LiquidAray() As String  'An array for the liquids
Dim FluidAray() As String    'An array to house wetting fluids
Dim fliudIndex As Integer    'this variable indexes the FluidAray for all forms to access the correct position of the array from any form

Private Sub Command1_Click()

Dim a$
Dim i As Integer
Dim ckSur, ckVisc, ckConv As Boolean

ckVisc = IsNumeric(txtVisc.Text)
ckSur = IsNumeric(txtSur.Text)
ckConv = IsNumeric(txtConvFactor.Text)

If targetIsLiqid = True Then
    ' they selected a liquid
    If NeedLiquid Then
        ' this liquid is for use in permeability
        ' so only the viscosity is important
        If txtVisc.Text = "" Then
            ' if no viscosity entered, use the value from the database
            txtVisc.Text = LiquidAray(cboLiqid.ListIndex + 1, 3)
            txtSur.Text = LiquidAray(cboLiqid.ListIndex + 1, 4)
        ElseIf ckVisc = False Then
            ' if a viscosity entered, give error if it is not a number
            numMess
            Exit Sub
        Else
            ' a viscosity was entered, and is a number, so update the database
            ' with the entered values.
            LiquidAray(cboLiqid.ListIndex + 1, 3) = txtVisc.Text
            LiquidAray(cboLiqid.ListIndex + 1, 4) = txtSur.Text
        End If
    End If
    If NeedFluid Then
        ' this liquid is for use as a wetting fluid
        ' so only the surface tension is important
        If txtSur.Text = "" Then
            ' if no surface tension entered, use the value from the database
            txtVisc.Text = FluidAray(cboLiqid.ListIndex + 1, 3)
            txtSur.Text = FluidAray(cboLiqid.ListIndex + 1, 4)
        ElseIf ckSur = False Then
            ' if a surface tension entered, give error if it is not a number
            numMess
            Exit Sub
        Else
            ' a surface tension was entered, and is a number, so update the database
            ' with the entered values.
            FluidAray(cboLiqid.ListIndex + 1, 3) = txtVisc.Text
            FluidAray(cboLiqid.ListIndex + 1, 4) = txtSur.Text
        End If
    End If
Else ' the target must be a gas
    ' gasses are only used for permeability, so only the viscosity is important
    If txtVisc.Text = "" Then
        ' if no viscosity entered, use the value from the database
        txtVisc.Text = GasAray(cboGas.ListIndex + 1, 3)
    ElseIf ckVisc = False Then
        ' if a viscosity is entered, and it is not a number, give error message
        numMess
        Exit Sub
    Else
        ' a viscosity was entered, and it is a number, so update database
        GasAray(cboGas.ListIndex + 1, 3) = txtVisc.Text
    End If
    
    If txtConvFactor.Text = "" Then
        txtConvFactor.Text = GasAray(cboGas.ListIndex + 1, 5)
    ElseIf Not ckConv Then
        numMess
        Exit Sub
    Else
        GasAray(cboGas.ListIndex + 1, 3) = txtConvFactor.Text
    End If
End If

    a$ = txtNome.Text
    selected_gas_or_liquid.type = epgl_other
    selected_gas_or_liquid.name = a$
    If NeedFluid Then
        ' only surface tension is importnat
        selected_gas_or_liquid.surfTen = myVal(txtSur.Text)
    ElseIf NeedGas And (targetIsLiqid = False) Then
        selected_gas_or_liquid.convFactor = GasAray(cboGas.ListIndex + 1, 5)
        gasflowconversionfactor = GasAray(cboGas.ListIndex + 1, 5)
        WPPS "main", "gas1", cboGas.Text, EXE_Path$ + "gasflowconversion.ini"
        WPPS "main", "conversion1", gasflowconversionfactor, EXE_Path$ + "gasflowconversion.ini"
        
        ' only viscosity is important
        ' check to see if this is a default value
        If LCase$(a$) = "water" Then
            selected_gas_or_liquid.type = epgl_water
        ElseIf LCase$(a$) = "air" Then
            selected_gas_or_liquid.type = epgl_air
        ElseIf LCase$(a$) = "argon" Then
            selected_gas_or_liquid.type = epgl_argon
        ElseIf LCase$(a$) = "helium" Then
            selected_gas_or_liquid.type = epgl_helium
        ElseIf LCase$(a$) = "nitrogen" Then
            selected_gas_or_liquid.type = epgl_nitrogen
        End If
        If selected_gas_or_liquid.type <> epgl_other Then
            complete_pgl selected_gas_or_liquid
            ' if the viscosity equals the value they selected
            ' then this can remain one of the default types
            ' Otherwise, it should be set back to the "other"
            ' type with their entered viscosity
            If selected_gas_or_liquid.viscosity <> myVal(txtVisc.Text) Then
                selected_gas_or_liquid.type = epgl_other
                selected_gas_or_liquid.viscosity = myVal(txtVisc.Text)
                selected_gas_or_liquid.name = a$
            End If
        Else
            ' it is indeed an "other"
            selected_gas_or_liquid.viscosity = myVal(txtVisc.Text)
        End If
    Else
        selected_gas_or_liquid.name = myVal(Me.txtVisc)
        selected_gas_or_liquid.name = Me.txtNome
        
    End If
    numGL = 0
    countArray
    Open EXE_Path$ + "GasLiq.txt" For Output As #1               '"C:\source_code\capwin\"
    Write #1, numGL
    For i = 1 To lenLiq                                                  'cboLiqid.ListCount
        Write #1, LiquidAray(i, 1), LiquidAray(i, 2), LiquidAray(i, 3), LiquidAray(i, 4)
    Next i
    For i = 1 To lenFl
        Write #1, FluidAray(i, 1), FluidAray(i, 2), FluidAray(i, 3), FluidAray(i, 4)
    Next i
    For i = 1 To lenGas
        Write #1, GasAray(i, 1), GasAray(i, 2), GasAray(i, 3), GasAray(i, 4), GasAray(i, 5)
    Next i
    Close #1
    Unload Me

End Sub
Private Sub countArray()
Dim i As Integer
 lenLiq = 0
 For i = 1 To UBound(LiquidAray, 1)
  If LiquidAray(i, 1) <> "" Then
   numGL = numGL + 1
   lenLiq = lenLiq + 1
  End If
 Next i
 lenGas = 0
 For i = 1 To UBound(GasAray, 1)
  If GasAray(i, 1) <> "" Then
   numGL = numGL + 1
   lenGas = lenGas + 1
  End If
 Next i
 lenFl = 0
 For i = 1 To UBound(FluidAray, 1)
  If FluidAray(i, 1) <> "" Then
   numGL = numGL + 1
   lenFl = lenFl + 1
  End If
 Next i
End Sub

Private Sub Command2_Click()
' if clear is true, this means we are in the middle of possibly adding a new gas or liquid
' and the "cancel" button is for clearing this adding, not for closing the form
 If clear = False Then
    ' we want to close the form and leave everything the way it originally was
    ' since selected_gas_or_liquid is both the return value and the initial value,
    ' and we don't modify it until they have properly selected something and are
    ' leaving, it should still hold what it originally held
    ' so we don't need to do anything special
    Unload Me
Else
    txtSur.Text = ""
    txtNome.Text = ""
    txtVisc.Text = ""
    txtConvFactor.Text = ""
    clear = False
    addGL = False
    If NeedGas Then
        cboGas.Enabled = True
        cmdGas.Enabled = True
    End If
    If NeedLiquid Or NeedFluid Then
        cboLiqid.Enabled = True
        cmdLiqid.Enabled = True
    End If
    txtNome.Text = ""           'clear all the text boxes
    txtSur.Text = ""
    txtVisc.Text = ""
    txtNome.Locked = True
    If targetIsLiqid Then
        cboLiqid.ListIndex = 0
        cboLiqid_Change
    Else
        cboGas.ListIndex = 0
        cboGas_Change
    End If
    Command1.Enabled = True
End If
End Sub
Private Sub Form_Load()
Dim found_it As Boolean
Dim i As Integer
Dim j As Integer
Dim numOfFields As Integer
Dim gasIndex As Integer
Dim liqidIndex As Integer

LoadTextStrings

'***********************************************************************************
'New code section
'Written by Edward Corvinelli
'10-17-05
'Purpose: To make it easier for the end user to select and/or edit a gas or liquid for any test
'that they may run. The modifications to the form will allow the user to add a gas or liquid
'if they wish. This first section will fill an array from a file with the names, viscosities and
'surface tensions of liquids and gases.
' Modified 3-13-06 by rvw to add comments, clean stuff up, make the calling and return values
' more consistant and get rid of unnecessary variables and globals
Dim cat, name, visc, surTen, convFactor, lname, gname As String   'Declare variable names for catagory,name,viscosity and
                                                                    'surface tension
Dim inputStr As String
Dim splitStr() As String

clear = False  ' we are not in the middle of allowing them to enter a new gas or liquid
If NeedGas = False Then
    cboGas.Enabled = False
    cmdGas.Enabled = False
End If
If NeedFluid = False And NeedLiquid = False Then
    cboLiqid.Enabled = False
    cmdLiqid.Enabled = False
End If
txtSur.Text = ""
txtNome.Text = ""
txtVisc.Text = ""
txtSur.Enabled = NeedFluid
txtVisc.Enabled = (Not NeedFluid)
addGL = False
gasNum = 0
liqNum = 0
fluNum = 0
found_it = False
gasIndex = -1
liqidIndex = -1
If Dir(EXE_Path$ + "GasLiq.txt") = "" Then
    useDefaultArray
Else
    Open EXE_Path$ + "GasLiq.txt" For Input As #1      'Open a file for input "C:\source_code\capwin\"
    Input #1, numGL                         ' get the first row of the file which contains the number of entries
    ReDim dump(1 To numGL, 1 To 5)    'dimension the dump array
    For i = 1 To numGL                      'Loop through the file
        Line Input #1, inputStr
        splitStr = split(inputStr, ",")
        numOfFields = UBound(splitStr)
        For j = 0 To numOfFields
            dump(i, j + 1) = Replace(splitStr(j), Chr(34), "")
        Next
        
        'Input #1, cat, name, visc, surTen  'read the file line and assign each variable a value
        
        'dump(i, 1) = cat                   'set array to the value of the variables
        'dump(i, 2) = name
        'dump(i, 3) = visc
        'dump(i, 4) = surTen
    Next i                      'go to next line in the file until the value in the header is reached
    Close #1                   'close file
End If
For i = 1 To numGL        'loop through the dump array and determine # of liquids qnd fluids
  If dump(i, 1) = "G" Then
    gasNum = gasNum + 1
  ElseIf dump(i, 1) = "L" Then
    liqNum = liqNum + 1
  Else
    fluNum = fluNum + 1
  End If
Next i
liqNum = liqNum + 50
gasNum = gasNum + 50
fluNum = fluNum + 50
ReDim GasAray(1 To gasNum, 1 To 5)
ReDim FluidAray(1 To fluNum, 1 To 4)
ReDim LiquidAray(1 To liqNum, 1 To 4)
indexL = 1                         'set the index to 1
indexG = 1                         'set the index to 1
indexF = 1
For i = 1 To numGL                 'Start to loop through the dump array
    If dump(i, 1) = "G" Then        'find the gasses through their catagory or cat from above
        GasAray(indexG, 1) = dump(i, 1) 'load the gas array with the cat
        GasAray(indexG, 2) = dump(i, 2) 'load the gas name
        GasAray(indexG, 3) = dump(i, 3) 'load the viscosity
        GasAray(indexG, 4) = dump(i, 4) 'load the surface tension if not applicable " " is put in
        GasAray(indexG, 5) = IIf(dump(i, 5) = "", 1, dump(i, 5))
        indexG = indexG + 1         'set the index for the next gas
    ElseIf dump(i, 1) = "F" Then
        FluidAray(indexF, 1) = dump(i, 1) 'load the fluid array cat
        FluidAray(indexF, 2) = dump(i, 2)   'load the fluis arrat name
        FluidAray(indexF, 3) = dump(i, 3)  'load the fluid array viscasity
        FluidAray(indexF, 4) = dump(i, 4) 'load the fluis array surface tension
        indexF = indexF + 1
     Else
        LiquidAray(indexL, 1) = dump(i, 1)  'load the liquid array with the cat
        LiquidAray(indexL, 2) = dump(i, 2)  'load the liquid array with the liquid name
        LiquidAray(indexL, 3) = dump(i, 3)  'load the liquid array with the liquid viscosity
        LiquidAray(indexL, 4) = dump(i, 4)  'load the liquid array with the liquid surface tension
        indexL = indexL + 1             'set the index for the next liquid
     End If
Next i
Dim tempStr As String
If NeedGas Then
    tempStr = ""
    For i = 1 To indexG - 1                   'set the control for the loop to that will
        tempStr = GasAray(i, 2)
        cboGas.AddItem tempStr              'add gas names to the drop down list
        'cboGas.ItemData(cboGas.NewIndex) = i
        If Not found_it Then
            If tempStr = selected_gas_or_liquid.name Then
                found_it = True
                gasIndex = cboGas.NewIndex
            End If
        End If
        tempStr = ""
    Next i                                  'Next in the loop
End If
tempStr = ""
If NeedLiquid Then
    For i = 1 To indexL - 1                   'set up control loop for the liquids
        tempStr = LiquidAray(i, 2)
        cboLiqid.AddItem tempStr             'add item to the liquids list
        'cboLiqid.ItemData(cboLiqid.NewIndex) = i
        If Not found_it Then
            If tempStr = selected_gas_or_liquid.name Then
                found_it = True
                liqidIndex = cboLiqid.NewIndex
            End If
        End If
    Next i
End If
If NeedFluid Then
    For i = 1 To indexF - 1                   'set up control loop for the liquids
        tempStr = FluidAray(i, 2)
        cboLiqid.AddItem tempStr             'add item to the liquids list
        'cboLiqid.ItemData(cboLiqid.NewIndex) = i
        If Not found_it Then
            If tempStr = selected_gas_or_liquid.name Then
                found_it = True
                liqidIndex = cboLiqid.NewIndex
            End If
        End If
    Next i
End If

If gasIndex >= 0 Then
    ' it is a gas, which is always for viscosity
    If NeedLiquid Then cboLiqid.ListIndex = -1 ' set liquid list to unselected
    cboGas.ListIndex = gasIndex ' set gas list to current element
    txtVisc.Text = selected_gas_or_liquid.viscosity
    targetIsLiqid = False
ElseIf liqidIndex >= 0 Then
    ' it is a liquid, which can be for either
    If NeedGas Then cboGas.ListIndex = -1 ' set gas list to unselected
    cboLiqid.ListIndex = liqidIndex ' set liquid list to current element
    If NeedFluid Then
         selected_gas_or_liquid.surfTen = txtSur.Text
    Else
         selected_gas_or_liquid.viscosity = txtVisc.Text
    End If
    targetIsLiqid = True
Else
    If selected_gas_or_liquid.type = epgl_undefined Then
        ' default to first item on list
        If NeedGas Then
            If NeedLiquid Then cboLiqid.ListIndex = -1
            cboGas.ListIndex = 0
            txtVisc.Text = GasAray(1, 3)
            targetIsLiqid = False
        Else
            cboLiqid.ListIndex = 0
            If NeedFluid Then
                txtSur.Text = FluidAray(1, 4)
            Else
                txtVisc.Text = LiquidAray(1, 3)
            End If
            targetIsLiqid = True
        End If
    Else
        ' try to add the initial selection to the list
        If NeedFluid Then
            ' it must be a fluid
            FluidAray(indexF, 1) = "F"
            FluidAray(indexF, 2) = selected_gas_or_liquid.name
            FluidAray(indexF, 3) = ""
            FluidAray(indexF, 4) = Format$(selected_gas_or_liquid.surfTen)
            cboLiqid.AddItem selected_gas_or_liquid.name
            'cboLiqid.ItemData(cboLiqid.NewIndex) = indexF
            cboLiqid.ListIndex = cboLiqid.NewIndex
            indexF = indexF + 1
            txtNome.Text = selected_gas_or_liquid.name
            txtSur.Text = Format$(selected_gas_or_liquid.surfTen)
            targetIsLiqid = True
        ElseIf NeedGas = False Or (NeedLiquid = True And selected_gas_or_liquid.viscosity >= 0.25) Then
            ' it must be a liquid or we quess that it is a liquid because of high viscosity
            LiquidAray(indexL, 1) = "L"
            LiquidAray(indexL, 2) = selected_gas_or_liquid.name
            LiquidAray(indexL, 3) = Format$(selected_gas_or_liquid.viscosity)
            LiquidAray(indexL, 4) = ""
            cboLiqid.AddItem selected_gas_or_liquid.name
            'cboLiqid.ItemData(cboLiqid.NewIndex) = indexF
            cboLiqid.ListIndex = cboLiqid.NewIndex
            indexL = indexL + 1
            txtNome.Text = selected_gas_or_liquid.name
            txtVisc.Text = Format$(selected_gas_or_liquid.surfTen)
            targetIsLiqid = True
        Else
            ' it must be a gas or we quess that it is a gas because of low viscosity
            GasAray(indexG, 1) = "G"
            GasAray(indexG, 2) = selected_gas_or_liquid.name
            GasAray(indexG, 3) = Format$(selected_gas_or_liquid.viscosity)
            GasAray(indexG, 4) = ""
            cboGas.AddItem selected_gas_or_liquid.name
            'cboGas.ItemData(cboLiqid.NewIndex) = indexF
            cboGas.ListIndex = cboGas.NewIndex
            indexG = indexG + 1
            txtNome.Text = selected_gas_or_liquid.name
            txtVisc.Text = Format$(selected_gas_or_liquid.surfTen)
            targetIsLiqid = False
        End If
    End If
End If
'edc 12-11-06 alter border color and caption
Me.Caption = Me.Caption & "    " & SubCaption
Me.BackColor = lngBorderColor
End Sub
Private Sub useDefaultArray()
Dim i As Integer
Open EXE_Path$ + "GasLiq.txt" For Output As #1
numGL = 8
ReDim dump(1 To 8, 1 To 5)
dump(1, 1) = "L"
dump(1, 2) = "Water"
dump(1, 3) = "1.0"
dump(1, 4) = "72.8"
dump(2, 1) = "F"
dump(2, 2) = "Silwick"
dump(2, 3) = "9.29853023"
dump(2, 4) = "20.1"
dump(3, 1) = "F"
dump(3, 2) = "Galwick"
dump(3, 3) = "8.0"
dump(3, 4) = "15.9"
dump(4, 1) = "F"
dump(4, 2) = "Porewick"
dump(4, 3) = "3.872"
dump(4, 4) = "16"
dump(5, 1) = "G"
dump(5, 2) = "Air"
dump(5, 3) = "0.018486176290633572"
dump(5, 4) = ""
dump(5, 5) = "1.0015"
dump(6, 1) = "G"
dump(6, 2) = "Helium"
dump(6, 3) = "0.0025"
dump(6, 4) = ""
dump(6, 5) = "1.4005"
dump(7, 1) = "G"
dump(7, 2) = "Nitrogen"
dump(7, 3) = "0.01756255075286845"
dump(7, 4) = ""
dump(7, 5) = "1.0000"
dump(8, 1) = "G"
dump(8, 2) = "Argon"
dump(8, 3) = "0.02217"
dump(8, 4) = ""
dump(8, 5) = "1.4047"
Write #1, numGL
For i = 1 To 8
    Write #1, dump(i, 1), dump(i, 2), dump(i, 3), dump(i, 4)
Next i
Close #1
End Sub
Private Sub cboGas_Click()
    Dim boxtext As String
    Dim i As Integer
    
    If cboGas.ListIndex = -1 Then Exit Sub
    targetIsLiqid = False
    
    txtNome.Text = cboGas.Text
    i = cboGas.ListIndex + 1
    txtVisc.Text = GasAray(i, 3)
    txtConvFactor = GasAray(i, 5)
    If NeedLiquid Then cboLiqid.ListIndex = -1
End Sub
Private Sub cboLiqid_Click()
Dim c As String
Dim i As Integer
If cboLiqid.ListIndex = -1 Then Exit Sub
targetIsLiqid = True
txtNome.Text = cboLiqid.Text
If NeedFluid Then
    i = cboLiqid.ListIndex + 1
    txtVisc.Text = FluidAray(i, 3)
    txtSur.Text = FluidAray(i, 4)
Else
    i = cboLiqid.ListIndex + 1
    txtVisc.Text = LiquidAray(i, 3)
    txtSur.Text = LiquidAray(i, 4)
End If
If NeedGas Then cboGas.ListIndex = -1
End Sub
'the purpose of this sub is to alter the objects on the form so the user may enter the add mode
'in order to create a new gas
Private Sub cmdGas_Click()
targetIsLiqid = False
check_addGL                 'when this sub is called it will toggel the form into or uot of add mode
'by checking the boolean value addGL if it is false it will assume that the user wants to add a gas and change
'thing accordingly
txtSur.Text = ""            'clear text box
If addGL = True Then
    txtNome.Locked = False      'unlock the name text box so that a name can be entered by the user
Else
    txtNome.Locked = True       'lock the name textbox to bar users from altering the contents
End If
If gasNum <= cboLiqid.ListCount + 1 Then
    probChild
End If
End Sub
Private Sub probChild()
MsgBox "This form has experienced a unexpected exception and must close. Reopen the form to add more liquids or gases." & _
                " Sorry for the inconveince", vbOKOnly, "Warning! Unexpected Exception..."
End Sub
'the purpose of this sub is to alter the objects on the form so the user may enter the add mode
'in order to create a new liquid
Private Sub cmdLiqid_Click()
targetIsLiqid = True
check_addGL             'when this sub is called it will toggel the form into or uot of add mode
'by checking the boolean value addGL if it is false it will assume that the user wants to add a gas and change
'thing accordingly
If addGL = True Then
    txtNome.Locked = False      'unlock the name text box so that a name can be entered by the user
    If NeedLiquid Then
        If liqNum <= cboLiqid.ListCount + 1 Then
            probChild
        End If
    Else
        If fluNum <= cboLiqid.ListCount + 1 Then
            probChild
        End If
    End If
Else
    txtNome.Locked = True       'lock the name textbox to bar users from altering the contents
End If
End Sub
Private Sub check_addGL()
If addGL = False Then       'check the status of the boolean
    clear = True
    addGL = True                'change it to true
    cboGas.Enabled = False      'disable the gas combo box
    cboLiqid.Enabled = False    'disable the liquid combo box
    cmdGas.Enabled = False
    cmdLiqid.Enabled = False
    txtNome.Text = ""           'clear all the text boxes
    txtSur.Text = ""
    txtVisc.Text = ""
    Command1.Enabled = False    'disable the ok button
    txtNome.Locked = False
    txtNome.SetFocus
End If
End Sub
Private Sub numMess()
 MsgBox "The value entered in either the Viscosity or  the Surface Tension fields is not a" & _
                " numerical value.Please enter a valid amount", vbOKOnly, "Warning! Numeric Exception..."
                txtVisc.Text = ""
                txtSur.Text = ""
                txtVisc.SetFocus
End Sub
Private Sub blankField()
 MsgBox "You have left a required field blank. Please enter a valid amount.", vbOKOnly, "Warning! Numeric Exception..."
                If txtNome.Text = "" Then
                    txtNome.SetFocus
                ElseIf txtVisc.Text = "" And txtVisc.Enabled = True Then
                    txtVisc.SetFocus
                Else: txtSur.SetFocus
                End If
End Sub
Private Sub ckFill()
If txtNome.Text = "" Then
    blankField
ElseIf NeedFluid Then
    If txtSur.Text = "" Then
        blankField
    End If
Else
    If txtVisc.Text = "" Then
            blankField
    End If
End If
End Sub
'This sub will execute when the Commit Changes button is clicked. It will determine if the user is trying to
'add a gas or liquid to the lists or just editing a
Private Sub cmdComit_Click()
'befor anything else check to see if the text in the viscosity and surface tension boxes is a number.
ckFill
If NeedFluid Then
    If IsNumeric(txtSur.Text) = False Then
        numMess
        Exit Sub
    End If
Else
    If IsNumeric(txtVisc.Text) = False Then
        numMess
        Exit Sub
    End If
End If
If addGL = True Then                                        'check the boolean for the adding a gas or liquid
    If targetIsLiqid = True Then                            'check the boolean for adding a liquid
        If NeedLiquid Then
            LiquidAray(indexL, 1) = "L"                         'set the first element to L for liquid
            LiquidAray(indexL, 2) = txtNome.Text                'set the next element  to the text of the name box
            LiquidAray(indexL, 3) = txtVisc.Text                'set the next array element to the viscosity box
            LiquidAray(indexL, 4) = txtSur.Text                 'set the next element to the surface tension text
            cboLiqid.AddItem txtNome.Text                           'display the name of the new liquid in the combobox
            cboLiqid.ListIndex = cboLiqid.ListCount - 1                               'inrect indexL to reflect the number of items in the array
            cboLiqid.Enabled = True
            Command1.Enabled = True
            txtNome.Locked = True
            indexL = indexL + 1
        Else
            FluidAray(indexF, 1) = "F"                         'set the first element to L for liquid
            FluidAray(indexF, 2) = txtNome.Text                'set the next element  to the text of the name box
            FluidAray(indexF, 3) = txtVisc.Text                'set the next array element to the viscosity box
            FluidAray(indexF, 4) = txtSur.Text                 'set the next element to the surface tension text
            cboLiqid.AddItem txtNome.Text                           'display the name of the new liquid in the combobox
            cboLiqid.ListIndex = cboLiqid.ListCount - 1                               'inrect indexL to reflect the number of items in the array
            cboLiqid.Enabled = True
            Command1.Enabled = True
            txtNome.Locked = True
            indexF = indexF + 1
        End If
    Else
            GasAray(indexG, 1) = "G"                                'Set the first element to G for Gas
            GasAray(indexG, 2) = txtNome.Text                       'set the second elelment to the name text box
            GasAray(indexG, 3) = txtVisc.Text                       'set the third element to the viscosity text box
            GasAray(indexG, 4) = ""                                 'set the fourth element to an empty set
            GasAray(indexG, 5) = txtConvFactor.Text
            cboGas.AddItem txtNome.Text                             'display the name of the new Gas in the combo box
            cboGas.Enabled = True
            Command1.Enabled = True
            txtNome.Locked = True
            cboGas.ListIndex = cboGas.ListCount - 1
            indexG = indexG + 1
    End If
Else      'AddGL is false hense the liquid or gas is just being edited
    If targetIsLiqid = True Then                            'liquid is being edited
        If NeedLiquid Then
            LiquidAray(cboLiqid.ListIndex + 1, 3) = txtVisc.Text   'Liquid array element 3 is reset to the contents of the viscosity textbox
            LiquidAray(cboLiqid.ListIndex + 1, 4) = txtSur.Text    'Liquid array element 4 is reset to the contents of the surface tension textbox
        Else
            FluidAray(cboLiqid.ListIndex + 1, 3) = txtVisc.Text   'Fluid array element 3 is reset to the contents of the viscosity textbox
            FluidAray(cboLiqid.ListIndex + 1, 4) = txtSur.Text    'Fluid array element 4 is reset to the contents of the surface tension textbox
        End If
    Else                                                    'Gas is being edited
        GasAray(cboGas.ListIndex + 1, 3) = txtVisc.Text           'Gas element 3 is rest to the contents of viscosity box
        GasAray(cboGas.ListIndex + 1, 4) = ""                     'Gas element 4 is reset to the value of empty set
        GasAray(cboGas.ListIndex + 1, 5) = txtConvFactor.Text
    End If
End If
addGL = False
End Sub
'This sub is called when the combo box value is changed and its main purpose is to update the
'various text boxes to reflect the users choise of gas or liquid
Private Sub cboGas_Change()
    targetIsLiqid = False
    txtNome.Text = ""                   'clear all text boxes of their current contents
    txtVisc.Text = ""
    txtSur.Text = ""
    txtConvFactor.Text = ""
    txtNome.Text = cboGas.Text          'Load the Gas name from the combo box
    txtVisc.Text = GasAray(cboGas.ListIndex + 1, 3) 'load the viscosity from the Gas array
    txtConvFactor.Text = GasAray(cboGas.ListIndex + 1, 3)
    txtSur.Text = ""
    If NeedLiquid Then cboLiqid.ListIndex = -1
End Sub

'This sub is to syncronize the the combo boxes and the text boxes to the same values so that when
'the user makes a choise in the combox the text boxes display the correct values for that luiquid.
Private Sub cboLiqid_Change()
targetIsLiqid = True
txtNome.Text = ""               'clears all text boxes of their previous contents
txtVisc.Text = ""
txtSur.Text = ""
txtNome.Text = cboLiqid.Text                    'sets text box to match the name of the combo box
If NeedLiquid Then
    txtVisc.Text = LiquidAray(cboLiqid.ListIndex + 1, 3)    'displays the fluid viscosity
    txtSur.Text = LiquidAray(cboLiqid.ListIndex + 1, 4)  'displays the suface tension of the liquid
Else
    txtVisc.Text = FluidAray(cboLiqid.ListIndex + 1, 3)    'displays the fluid viscosity
    txtSur.Text = FluidAray(cboLiqid.ListIndex + 1, 4)  'displays the suface tension of the liquid
End If
If NeedGas Then cboGas.ListIndex = -1
End Sub

Public Sub LoadTextStrings()
    ' Load text elements for this form from external .ini file
    
    Dim i As Integer
    
    ' Form elements
    GLSel.Caption = gpps2("glsel", "window title", language$, GLSel.Caption)
    set_fontstuff txtNome, default_font
    Command1.Caption = gpps2("glsel", "command1", language$, Command1.Caption)
    set_fontname Command1, default_font
    Command2.Caption = gpps2("glsel", "command2", language$, Command2.Caption)
    set_fontname Command2, default_font
    
    ' Other text
    ts$(1) = gpps2("glsel", "ts1", language$, "Error: incorrect viscosity value")
    ts$(2) = gpps2("glsel", "ts2", language$, "Error: incorrect gas/liquid name value")
    ts$(3) = gpps2("glsel", "ts3", language$, "Unit")
    
End Sub

