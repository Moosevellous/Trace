Attribute VB_Name = "LogFunctions"
Public Function SPLSUM(ParamArray rng1() As Variant) As Variant
On Error Resume Next

Dim C As Range
Dim i As Long

SPLSUM = -99
For i = LBound(rng1) To UBound(rng1)
'Debug.Print TypeName(Rng1(i))
    If TypeName(rng1(i)) = "Double" Then
        If rng1(i) > 0 Then 'negative values are ignored
        SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (rng1(i) / 10)))
        End If
    ElseIf TypeName(rng1(i)) = "Range" Then
        For Each C In rng1(i).Cells
            If C.Value <> Empty Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (C.Value / 10)))
            End If
        Next C
    End If
Next i

''check for negative
'If SPLSUM < -5 Then
'SPLSUM = ""
'End If

End Function

Public Function SPLAV(ParamArray rng1() As Variant) As Variant
On Error Resume Next

Dim C As Range
Dim i As Long
Dim n As Integer
SPLAV = -99
n = 0
For i = LBound(rng1) To UBound(rng1)
'Debug.Print TypeName(Rng1(i))
    If TypeName(rng1(i)) = "Double" Then
        If rng1(i) > 0 Then 'negative values are ignored
        SPLAV = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLAV / 10)) + (10 ^ (rng1(i) / 10)))
        n = n + 1
        End If
    ElseIf TypeName(rng1(i)) = "Range" Then
        For Each C In rng1(i).Cells
            If C.Value <> Empty And IsNumeric(C.Value) Then
            SPLAV = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLAV / 10)) + (10 ^ (C.Value / 10)))
            n = n + 1
            End If
        Next C
    End If
Next i

'Average +10log(1/n) in log domain
SPLAV = SPLAV + 10 * Application.WorksheetFunction.Log10(1 / n)

''check for negative
'If SPLSUM < -5 Then
'SPLSUM = ""
'End If

End Function

Public Function SPLMINUS(SPLtotal As Double, SPL2 As Double) As Variant
On Error Resume Next

SPLMINUS = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLtotal / 10)) - (10 ^ (SPL2 / 10)))

End Function

Public Function TL_ThirdsToOctave(rngInput As Range) As Variant

Dim isNegative As Boolean
Dim TL1 As Single, TL2 As Single, TL3 As Single

TL1 = rngInput.Cells(1, 1).Value
TL2 = rngInput.Cells(1, 2).Value
TL3 = rngInput.Cells(1, 3).Value

    'switch for pos/neg sign (losses should be positive)
    If TL1 < 0 And TL2 < 0 And TL3 < 0 Then isNegative = True

    'flip signs
    If TL1 < 0 Then TL1 = TL1 * -1
    If TL2 < 0 Then TL2 = TL2 * -1
    If TL3 < 0 Then TL3 = TL3 * -1
    

    If isNegative Then 'return result as negative
    TL_ThirdsToOctave = 10 * Application.WorksheetFunction.Log10((1 / 3) * ((10 ^ (-TL1 / 10)) + (10 ^ (-TL2 / 10)) + (10 ^ (-TL3 / 10))))
    Else 'return result as positive
    TL_ThirdsToOctave = -10 * Application.WorksheetFunction.Log10((1 / 3) * ((10 ^ (-TL1 / 10)) + (10 ^ (-TL2 / 10)) + (10 ^ (-TL3 / 10))))
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'TODO
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

'Function SPLSUMIF()

'Function SPLAVIF()


