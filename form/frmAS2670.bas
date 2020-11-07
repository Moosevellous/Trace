VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAS2670 
   Caption         =   "Insert AS 2670 Vibration Curve"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5115
   OleObjectBlob   =   "frmAS2670.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAS2670"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
btnOkPressed = False

' Position - centre of screen
With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
End With

' Populate drop down boxes
PopulatecBoxAxis
PopulatecBoxMult
PopulatecBoxOrder

' Populate default values
Me.cBoxAxis.Text = "Z"
Me.cBoxMult.Text = "Critical Working Areas"
Me.txtMult.Text = 1
Me.cBoxOrder.Text = "Acceleration (m/s/s)"
'chkRate.Value = False

' Default values for toggle variables
AS2670_dbUnit = False
AS2670_RateCurve = False

End Sub

Private Sub btnOK_Click()

btnOkPressed = True

' Assign value to variable for Axis drop down box
Select Case Me.cBoxAxis.Text
Case Is = "Z"
    SelectedAxis = "z"
Case Is = "XY"
    SelectedAxis = "xy"
Case Is = "Combined (XYZ)"
    SelectedAxis = "comb."
End Select

' Assign value to variable for Order drop down box
Select Case Me.cBoxOrder.Text
Case Is = "Acceleration (m/s/s)"
    SelectedOrder = "Accel"
Case Is = "Velocity (m/s)"
    SelectedOrder = "Vel"
End Select

' Assign values from form to public variables
AS2670_Axis = SelectedAxis
AS2670_Multiplier = Me.txtMult.Value
AS2670_Order = SelectedOrder

' Close and unload form
Me.Hide
Unload Me

End Sub

Private Sub btnCancel_Click()

btnOkPressed = False

' Close and unload form
Me.Hide
Unload Me

End Sub

Private Sub cBoxMult_Change()
    
    Select Case Me.cBoxMult.Text
    Case Is = "Critical Working Areas"
        Me.txtMult.Value = 1
    Case Is = "Residential - Night"
        Me.txtMult.Value = 1.4
    Case Is = "Residential - Day"
        Me.txtMult.Value = 2
    Case Is = "Office"
        Me.txtMult.Value = 4
    Case Is = "Workshop"
        Me.txtMult.Value = 8
    End Select
    
End Sub

Private Sub optLinUnits_Click()
    AS2670_dbUnit = False
End Sub

Private Sub optdBUnits_Click()
    AS2670_dbUnit = True
End Sub

Sub PopulatecBoxAxis()

    With Me.cBoxAxis
        .AddItem ("Z")
        .AddItem ("XY")
        .AddItem ("Combined (XYZ)")
    End With

End Sub

Sub PopulatecBoxMult()

    With Me.cBoxMult
        .AddItem ("Critical Working Areas")
        .AddItem ("Residential - Night")
        .AddItem ("Residential - Day")
        .AddItem ("Office")
        .AddItem ("Workshop")
    End With

End Sub

Sub PopulatecBoxOrder()

    With Me.cBoxOrder
        .AddItem ("Acceleration (m/s/s)")
        .AddItem ("Velocity (m/s)")
    End With

End Sub

Private Sub optRateUserInput_Click()
    AS2670_RateCurve = False
    cBoxMult.Enabled = True
    cBoxMult.BackColor = &HC0FFFF
    txtMult.Enabled = True
    
    lblPlace.ForeColor = &H80000012
    lblMultiplier.ForeColor = &H80000012
    lblEquals.ForeColor = &H80000012
End Sub

Private Sub optRateExisting_Click()
    AS2670_RateCurve = True
    cBoxMult.Enabled = False
    cBoxMult.BackColor = &H8000000F
    txtMult.Enabled = False
    
    lblPlace.ForeColor = &H80000006
    lblMultiplier.ForeColor = &H80000006
    lblEquals.ForeColor = &H80000006
End Sub

'Private Sub chkRate_Click()
'
'    If chkRate.Value = False Then
'        AS2670_RateCurve = False
'        cBoxMult.Enabled = True
'        cBoxMult.BackColor = &HC0FFFF
'        txtMult.Enabled = True
'
'        lblPlace.ForeColor = &H80000012
'        lblMultiplier.ForeColor = &H80000012
'        lblEquals.ForeColor = &H80000012
'
'    Else
'        AS2670_RateCurve = True
'        cBoxMult.Enabled = False
'        cBoxMult.BackColor = &H8000000F
'        txtMult.Enabled = False
'
'        lblPlace.ForeColor = &H80000006
'        lblMultiplier.ForeColor = &H80000006
'        lblEquals.ForeColor = &H80000006
'
'    End If
'
'End Sub


