VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAS2670 
   Caption         =   "Insert AS 2670 Vibration Curve"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12090
   OleObjectBlob   =   "frmAS2670.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAS2670"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnHelp_Click()
GotoWikiPage ("Vibration#as2670")
End Sub


Private Sub cBoxAxis_Change()
PreviewValues
SelectNewPicture
End Sub

Private Sub cBoxOrder_Change()
PreviewValues
SelectNewPicture
End Sub

Sub SelectNewPicture()

Dim ImagePath As String

ImagePath = "img\AS2670_"
    
    'select axis
    Select Case Me.cBoxAxis.ListIndex
    Case 0 'Z
    ImagePath = ImagePath + "Z"
    Case 1 'XY
    ImagePath = ImagePath + "XY"
    Case 2 'Combined
    ImagePath = ImagePath + "Combined"
    End Select
    
    'select Accel/Vel
    Select Case Me.cBoxOrder.ListIndex
    Case 0 'accel
    ImagePath = ImagePath + "Accel"
    Case 1 'vel (rms)
    ImagePath = ImagePath + "Vel"
    End Select

ImagePath = ImagePath + ".jpg"

UpdatePicture (ImagePath)

End Sub

Sub UpdatePicture(FilePath As String)
Dim PathStr As String
GetSettings
PathStr = ROOTPATH & "\" & FilePath
    If Dir(PathStr, vbNormal) <> "" Then
    Me.imgCurves.Picture = LoadPicture(ROOTPATH & "\" & FilePath)
    End If
End Sub

Private Sub UserForm_Activate()
btnOkPressed = False

' Position - centre of screen
With Me
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
End With
btnOkPressed = False

' Populate drop down boxes
PopulatecBoxAxis
PopulatecBoxMult
PopulatecBoxOrder

' Populate default values
Me.cBoxAxis.Text = "Z"
Me.cBoxMult.Text = "Critical Working Areas"
Me.txtMult.Text = 1
Me.cBoxOrder.Text = "Acceleration (m/s/s)"


' Default values for toggle variables
AS2670_dbUnit = False
AS2670_RateCurve = False

End Sub

Private Sub btnOK_Click()

Dim SelectedAxis As String
Dim SelectedOrder As String

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
PreviewValues
End Sub

Private Sub optLinUnits_Click()
AS2670_dbUnit = False
PreviewValues
End Sub

Private Sub optdBUnits_Click()
AS2670_dbUnit = True
PreviewValues
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
Me.cBoxMult.Enabled = True
Me.cBoxMult.BackColor = &HC0FFFF
Me.txtMult.Enabled = True
Me.RefVibRange.Enabled = False
Me.RefVibRange.BackColor = &H8000000F

Me.lblPlace.ForeColor = &H80000012
Me.lblMultiplier.ForeColor = &H80000012
Me.lblEquals.ForeColor = &H80000012

PreviewValues
End Sub

Private Sub optRateExisting_Click()
AS2670_RateCurve = True
Me.cBoxMult.Enabled = False
Me.cBoxMult.BackColor = &H8000000F
Me.txtMult.Enabled = False
Me.RefVibRange.Enabled = True
Me.RefVibRange.BackColor = &HC0FFFF

Me.lblPlace.ForeColor = &H80000006
Me.lblMultiplier.ForeColor = &H80000006
Me.lblEquals.ForeColor = &H80000006

PreviewValues
End Sub

Sub PreviewValues()

Dim Mode As String
Dim SelectedAxis As String
Dim SelectedOrder As String

    If Me.optdBUnits.Value = True Then
    Mode = "dB"
    Else
    Mode = "Linear"
    End If
    
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

'check for values
    If Me.cBoxAxis.Value <> "" And Me.cBoxOrder <> "" Then
    'preview
    Me.txt1.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 1, SelectedOrder, Mode)
    Me.txt1_25.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 1.25, SelectedOrder, Mode)
    Me.txt1_6.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 1.6, SelectedOrder, Mode)
    Me.txt2.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 2, SelectedOrder, Mode)
    Me.txt2_5.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 2.5, SelectedOrder, Mode)
    Me.txt3_15.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 3.15, SelectedOrder, Mode)
    Me.txt4.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 4, SelectedOrder, Mode)
    Me.txt5.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 5, SelectedOrder, Mode)
    Me.txt6.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 6.3, SelectedOrder, Mode)
    Me.txt8.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 8, SelectedOrder, Mode)
    Me.txt10.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 10, SelectedOrder, Mode)
    Me.txt12.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 12, SelectedOrder, Mode)
    Me.txt16.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 16, SelectedOrder, Mode)
    Me.txt20.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 20, SelectedOrder, Mode)
    Me.txt25.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 25, SelectedOrder, Mode)
    Me.txt31.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 31, SelectedOrder, Mode)
    Me.txt40.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 40, SelectedOrder, Mode)
    Me.txt50.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 50, SelectedOrder, Mode)
    Me.txt63.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 63, SelectedOrder, Mode)
    Me.txt80.Value = AS2670_Curve(SelectedAxis, Me.txtMult.Value, 80, SelectedOrder, Mode)
    End If
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


