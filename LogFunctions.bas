Attribute VB_Name = "LogFunctions"
Public Function SPLSUM(ParamArray Rng1() As Variant) As Variant
On Error Resume Next

Dim c As Range
Dim i As Long

SPLSUM = -99
For i = LBound(Rng1) To UBound(Rng1)
'Debug.Print TypeName(Rng1(i))
    If TypeName(Rng1(i)) = "Double" Then
        If Rng1(i) > 0 Then 'negative values are ignored
        SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (Rng1(i) / 10)))
        End If
    ElseIf TypeName(Rng1(i)) = "Range" Then
        For Each c In Rng1(i).Cells
            If c.Value <> Empty Then
            SPLSUM = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLSUM / 10)) + (10 ^ (c.Value / 10)))
            End If
        Next c
    End If
Next i

''check for negative
'If SPLSUM < 0 Then
'SPLSUM = ""
'End If
End Function

Public Function SPLAV(ParamArray Rng1() As Variant) As Variant
On Error Resume Next

Dim c As Range
Dim i As Long
Dim n As Integer
SPLAV = -99
n = 0
For i = LBound(Rng1) To UBound(Rng1)
'Debug.Print TypeName(Rng1(i))
    If TypeName(Rng1(i)) = "Double" Then
        If Rng1(i) > 0 Then 'negative values are ignored
        SPLAV = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLAV / 10)) + (10 ^ (Rng1(i) / 10)))
        n = n + 1
        End If
    ElseIf TypeName(Rng1(i)) = "Range" Then
        For Each c In Rng1(i).Cells
            If c.Value <> Empty Then
            SPLAV = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLAV / 10)) + (10 ^ (c.Value / 10)))
            n = n + 1
            End If
        Next c
    End If
Next i

'Average +10log(1/n) in log domain
SPLAV = SPLAV + 10 * Application.WorksheetFunction.Log10(1 / n)

''check for negative
'If SPLSAV < 0 Then
'SPLAV = ""
'End If

End Function

Public Function SPLMINUS(SPLtotal As Double, SPL2 As Double) As Variant
On Error Resume Next

SPLMINUS = 10 * Application.WorksheetFunction.Log10((10 ^ (SPLtotal / 10)) - (10 ^ (SPL2 / 10)))

End Function

