Attribute VB_Name = "Vibration"
Public VibRef As String
Public ConversionFactorStr As String
Public VibConversionDescription As String
Public BuildingType As String
Public AmplificationType As String
''''''''''
'FUNCTIONS
''''''''''


Function VcCurve(CurveName As String, freq As String, Optional Mode As String)

Dim VC_OR() As Variant
Dim VC_A() As Variant
Dim VC_B() As Variant
Dim VC_C() As Variant
Dim VC_D() As Variant
Dim VC_E() As Variant
'Dim ChosenCurve() As Variant

'frequencies
'2hz, 2.5hz, 3.15hz, 4hz, 5hz, 6.3hz, 8hz, 10hz, 12.5hz, 16hz, 20hz, 25hz, 31.5hz, 40hz, 50hz, 63hz, 80hz
VC_E = Array(0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032, 0.0032)
VC_D = Array(0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064, 0.0064)
VC_C = Array(0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013, 0.013)
VC_B = Array("-", "-", "-", 0.05, 0.0397, 0.0315, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025)
VC_A = Array("-", "-", "-", 0.102, 0.081, 0.0643, 0.051, 0.051, 0.051, 0.051, 0.051, 0.051, 0.051, 0.051, 0.051, 0.051, 0.051)
VC_OR = Array(0.306, 0.2548, 0.2122, 0.1767, 0.1471, 0.1225, 0.102, 0.102, 0.102, 0.102, 0.102, 0.102, 0.102, 0.102, 0.102, 0.102, 0.102) 'operating room
    
'Debug.Print CurveName

ChosenCurve = ""
    Select Case CurveName
    Case "VC-OR"
    ChosenCurve = VC_OR
    Case "VC-A"
    ChosenCurve = VC_A
    Case "VC-B"
    ChosenCurve = VC_B
    Case "VC-C"
    ChosenCurve = VC_C
    Case "VC-D"
    ChosenCurve = VC_D
    Case "VC-E"
    ChosenCurve = VC_E
    End Select

f = freqStr2Num(freq)

VcCurve = "-" 'catch for errors

    Select Case f
    Case Is = 2 'VC curves start at 2Hz
    VcCurve = ChosenCurve(0) 'arrays start at 0
    Case Is = 2.5
    VcCurve = ChosenCurve(1)
    Case Is = 3.15
    VcCurve = ChosenCurve(2)
    Case Is = 4
    VcCurve = ChosenCurve(3)
    Case Is = 5
    VcCurve = ChosenCurve(4)
    Case Is = 6.3
    VcCurve = ChosenCurve(5)
    Case Is = 8
    VcCurve = ChosenCurve(6)
    Case Is = 10
    VcCurve = ChosenCurve(7)
    Case Is = 12.5
    VcCurve = ChosenCurve(8)
    Case Is = 16
    VcCurve = ChosenCurve(9)
    Case Is = 20
    VcCurve = ChosenCurve(10)
    Case Is = 25
    VcCurve = ChosenCurve(11)
    Case Is = 31.5
    VcCurve = ChosenCurve(12)
    Case Is = 40
    VcCurve = ChosenCurve(13)
    Case Is = 50
    VcCurve = ChosenCurve(14)
    Case Is = 63
    VcCurve = ChosenCurve(15)
    Case Is = 80
    VcCurve = ChosenCurve(16)
    End Select

    If VcCurve <> "-" And Mode = "dB" Then
    VcCurve = 20 * Application.WorksheetFunction.Log10(VcCurve / 0.000001)
    End If


End Function


Function VcRate(DataTable As Variant, freqTable As Variant)

Dim MaxCurve As Integer
Dim CurrentCurve As Integer

MaxCurve = 0

MapValue = Array("VC-E", "VC-D", "VC-C", "VC-B", "VC-A", "VC-OR")
    
    For i = 0 To 26 '26 columns is all you'll need
    
'    Debug.Print "Value:" & freqTable(i)
'    Debug.Print "Value:" & DataTable(i)
    
        Select Case DataTable(i)
        
        Case Is > VcCurve("VC-OR", CStr(freqTable(i)))
        CurrentCurve = 6
        Case Is > VcCurve("VC-A", CStr(freqTable(i)))
        CurrentCurve = 5
        Case Is > VcCurve("VC-B", CStr(freqTable(i)))
        CurrentCurve = 4
        Case Is > VcCurve("VC-C", CStr(freqTable(i)))
        CurrentCurve = 3
        Case Is > VcCurve("VC-D", CStr(freqTable(i)))
        CurrentCurve = 2
        Case Is > VcCurve("VC-E", CStr(freqTable(i)))
        CurrentCurve = 1
        End Select
        
        If CurrentCurve > MaxCurve Then
        MaxCurve = CurrentCurve
        End If
        
    Next i

VcRate = MapValue(MaxCurve)

End Function



'''''''''''''''''''
'SUBS
'''''''''''''''''''

Sub VibLin2DB(SheetType As String)
CheckRow (Selection.Row)
frmVibUnits.Show
    If btnOkPressed = True Then
    Cells(Selection.Row, 2).Value = "Convert to dB"
    Cells(Selection.Row, 5).Value = "=20*LOG(" & Cells(Selection.Row - 1, 5).Address(False, False) & "/" & VibRef & ")"
    ExtendFunction (SheetType)
    End If
End Sub


Sub VibDB2Lin(SheetType As String)
CheckRow (Selection.Row)
frmVibUnits.Show
    If btnOkPressed = True Then
    Cells(Selection.Row, 2).Value = "Convert to Linear"
    Cells(Selection.Row, 5).Value = "=" & VibRef & "*10^(E" & Selection.Row - 1 & "/20)"
    ExtendFunction (SheetType)
    End If
End Sub


Sub CouplingLoss(SheetType As String)

'Dim CRL(19) As Integer
'Dim LargeMasonryOnPiles() As Integer
'Dim LargeMasonryOnSpreadFootings() As Integer
'Dim TwoToFourStoreyMasonryOnSpreadFootings() As Integer
'Dim OneToTwoStoreyCommercial() As Integer
'Dim SingleResidential() As Integer
'Dim SelectedLoss() As Integer
'set Coupling Loss values, one-third octave bands from 5Hz onwards

CRL = Array(2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 2, 2, 2)
LargeMasonryOnPiles = Array(6, 6, 6, 6, 7, 7, 7, 8, 9, 10, 11, 12, 13, 13, 14, 14, 15, 15, 15)
LargeMasonryOnSpreadFootings = Array(11, 11, 11, 11, 12, 13, 14, 14, 15, 15, 15, 15, 14, 14, 14, 14, 13, 12, 11)
TwoToFourStoreyMasonryOnSpreadFootings = Array(5, 6, 6, 7, 9, 11, 11, 12, 13, 13, 13, 13, 13, 12, 12, 11, 10, 9, 8)
OneToTwoStoreyCommercial = Array(4, 5, 5, 6, 7, 8, 8, 9, 9, 9, 9, 9, 9, 8, 8, 8, 7, 6, 5)
SingleResidential = Array(3, 3, 4, 4, 5, 5, 6, 6, 6, 6, 6, 6, 6, 5, 5, 5, 4, 4, 4)

frmCouplingLoss.Show

If btnOkPressed = False Then End

    If SheetType = "LF_TO" Then
    
        Select Case BuildingType 'public variable
        Case Is = "CRL"
        SelectedLoss = CRL
        Case Is = "Large Masonry On Piles"
        SelectedLoss = LargeMasonryOnPiles
        Case Is = "Large Masonry on Spread Footings"
        SelectedLoss = LargeMasonryOnSpreadFootings
        Case Is = "2-4 Storey Masonry on Spread Footings"
        SelectedLoss = TwoToFourStoreyMasonryOnSpreadFootings
        Case Is = "1-2 Storey Commercial"
        SelectedLoss = OneToTwoStoreyCommercial
        Case Is = "Single Residential"
        SelectedLoss = SingleResidential
        End Select
        
        If IsEmpty(SelectedLoss) Then End
        
        For i = LBound(SelectedLoss) To UBound(SelectedLoss)
        Cells(Selection.Row, 12 + i).Value = -1 * SelectedLoss(i) 'negative values!
        Next i
    
    Cells(Selection.Row, 2).Value = "Coupling Loss: " & BuildingType
    
    End If
    
End Sub

Sub BuildingAmplification(SheetType As String)

FloorVibration = Array(10, 10, 10, 10, 10, 10, 10, 11, 11, 11, 10, 9, 9, 0, 0, 0, 0, 0, 0)
GBN = Array(0, 0, 0, 0, 0, 0, 6, 7, 7, 7, 6, 6, 5, 5, 4, 3, 2, 1, 1)

frmBuildingAmplification.Show

If btnOkPressed = False Then End

    If SheetType = "LF_TO" Then
    
        Select Case AmplificationType 'public variable
        Case Is = "Ground-borne Noise"
        SelectedLoss = GBN
        Case Is = "Floor Vibration"
        SelectedLoss = FloorVibration
        End Select
        
        If IsEmpty(SelectedLoss) Then End
        
        For i = LBound(SelectedLoss) To UBound(SelectedLoss)
            If SelectedLoss(i) <> 0 Then
            Cells(Selection.Row, 12 + i).Value = SelectedLoss(i) 'negative values!
            End If
        Next i
    
    Cells(Selection.Row, 2).Value = "Building Amplification: " & AmplificationType
    Else
    ErrorLFTOOnly
    End If

End Sub

Sub PutVCcurve(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

msg = MsgBox("Linear values (mm/s)? " & chr(10) & "[Note that 'No' will choose dB mode.]", vbYesNoCancel, "Lin/Log mode")
If msg = vbCancel Then End

Call ParameterMerge(Selection.Row, SheetType)

    If Left(SheetType, 3) = "OCT" Then
    ErrorLFTOOnly
    ElseIf Left(SheetType, 2) = "TO" Then
    ErrorLFTOOnly
    ElseIf SheetType = "LF_TO" Then
    Cells(Selection.Row, 32) = "VC-A"
        If msg = vbYes Then
        Cells(Selection.Row, 5).Value = "=VCcurve($AF" & Selection.Row & ",E$6)"
        ElseIf msg = vbNo Then 'dB mode
        Cells(Selection.Row, 5).Value = "=VCcurve($AF" & Selection.Row & ",E$6,""dB"")"
        End If
    End If
    
Cells(Selection.Row, 2).Value = "VC Curve"
ExtendFunction (SheetType)

With Cells(Selection.Row, 32).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="VC-OR,VC-A,VC-B,VC-C,VC-D,VC-E"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
fmtUserInput SheetType, True
End Sub


Sub VibConvert(SheetType As String)
Dim FormulaStr As String

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
frmVibConvert.Show

If btnOkPressed = False Then End
    
    If SheetType = "LF_TO" Then
    FormulaStr = Replace(ConversionFactorStr, "pi", "PI()")
    FormulaStr = Replace(FormulaStr, "f", "E$6")
    FormulaStr = Replace(FormulaStr, chr(178), "^2")
    Cells(Selection.Row, 5).Value = "=" & FormulaStr
    ExtendFunction (SheetType)
    Range(Cells(Selection.Row, 5), Cells(Selection.Row, 31)).NumberFormat = "0E+0"
    Cells(Selection.Row, 2).Value = "Vibration Conversion"
    Cells(Selection.Row, 2).ClearComments
    Cells(Selection.Row, 2).AddComment (VibConversionDescription)
    Cells(Selection.Row, 2).Comment.Shape.width = 150
    Cells(Selection.Row, 2).Comment.Shape.height = 25
    'Range("C1").Comment.Shape.ScaleWidth 5.87, msoFalse, msoScaleFromTopLeft
    Else
    ErrorLFTOOnly
    End If

End Sub
