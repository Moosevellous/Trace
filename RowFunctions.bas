Attribute VB_Name = "RowFunctions"
Public UserSelectedAddress As String
Public SumOrAverage As String

Public Sub CheckRow(rw As Integer)
'Checks that user isn't in header rows. These rows are protected by this function. None shall Pass.
If rw <= 7 Then End
End Sub

Sub ClearRw(SheetType As String)
Dim rw As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Selection.Rows.Count > 1 Then
    msg = MsgBox("Are you sure you want to clear rows?", vbYesNo, "Check")
    Else
    msg = vbYes
    End If

    If msg = vbYes Then
        For rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
        Cells(rw, 2).ClearContents
        Cells(rw, 2).ClearComments
            If Left(SheetType, 3) = "OCT" Then
            Range(Cells(rw, 5), Cells(rw, 15)).ClearContents
            Range(Cells(rw, 2), Cells(rw, 15)).Font.ColorIndex = 0
            Range(Cells(rw, 2), Cells(rw, 15)).Interior.ColorIndex = 0 'no colour
            Cells(rw, 14).Validation.Delete 'for dropdown boxes
            Cells(rw, 15).Validation.Delete 'for dropdown boxes
            Range(Cells(rw, 14), Cells(rw, 15)).Merge
            Cells(rw, 14).ClearComments
            Range(Cells(rw, 3), Cells(rw, 13)).FormatConditions.Delete 'removes heatmap
            Cells(rw, 14).NumberFormat = "General"
            ElseIf Left(SheetType, 2) = "TO" Then
            Range(Cells(rw, 5), Cells(rw, 27)).ClearContents
            Range(Cells(rw, 2), Cells(rw, 27)).Font.ColorIndex = 0
            Range(Cells(rw, 2), Cells(rw, 27)).Interior.ColorIndex = 0 'no colour
            Cells(rw, 26).Validation.Delete 'for dropdown boxes
            Cells(rw, 27).Validation.Delete 'for dropdown boxes
            Range(Cells(rw, 26), Cells(rw, 27)).Merge
            Cells(rw, 26).ClearComments
            Range(Cells(rw, 3), Cells(rw, 27)).FormatConditions.Delete 'removes heatmap
            Cells(rw, 26).NumberFormat = "General"
            End If
        Next rw
    End If
Call ParameterMerge(Selection.Row, SheetType)
End Sub

Sub FlipSign(SheetType As String)
Dim rw As Integer
CheckRow (Selection.Row)

    For rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
        For Col = Selection.Column To Selection.Column + Selection.Columns.Count - 1
            If Cells(rw, Col).HasFormula Then
                If Mid(Cells(rw, Col).Formula, 2, 1) = "-" Then
                Cells(rw, Col).Formula = Replace(Cells(rw, Col).Formula, "=-", "=", 1, Len(Cells(rw, Col).Formula), vbTextCompare)
                Else
                Cells(rw, Col).Formula = Replace(Cells(rw, Col).Formula, "=", "=-", 1, Len(Cells(rw, Col).Formula), vbTextCompare)
                End If
            Else
                Cells(rw, Col).Value = Cells(rw, Col).Value * -1
            End If
        Next Col
    Next rw

End Sub

Sub A_weight_oct(SheetType As String)

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

If SheetType = "OCT" Then
    For Col = 2 To 13
        Select Case Col
        Case 2
        Cells(7, Col).Value = "A Weighting"
        Case 5
        Cells(7, Col).Value = -39.4
        Case 6
        Cells(7, Col).Value = -26.2
        Case 7
        Cells(7, Col).Value = -16.1
        Case 8
        Cells(7, Col).Value = -8.6
        Case 9
        Cells(7, Col).Value = -3.2
        Case 10
        Cells(7, Col).Value = 0#
        Case 11
        Cells(7, Col).Value = -1.2
        Case 12
        Cells(7, Col).Value = 1#
        Case 13
        Cells(7, Col).Value = -1.1
        End Select
    Next Col
ElseIf SheetType = "OCTA" Then
    For Col = 2 To 13
        Select Case Col
        Case 2
        Cells(7, Col).Value = "A Weighting"
        Case 5
        Cells(7, Col).Value = 39.4
        Case 6
        Cells(7, Col).Value = 26.2
        Case 7
        Cells(7, Col).Value = 16.1
        Case 8
        Cells(7, Col).Value = 8.6
        Case 9
        Cells(7, Col).Value = 3.2
        Case 10
        Cells(7, Col).Value = 0#
        Case 11
        Cells(7, Col).Value = 1.2
        Case 12
        Cells(7, Col).Value = 1#
        Case 13
        Cells(7, Col).Value = 1.1
        End Select
    Next Col
ElseIf SheetType = "TO" Then
 For Col = 2 To 25
        Select Case Col
        Case 2
        Cells(7, Col).Value = "A Weighting"
        Case 5
        Cells(7, Col).Value = -30.2
        Case 6
        Cells(7, Col).Value = -26.2
        Case 7
        Cells(7, Col).Value = -22.5
        Case 8
        Cells(7, Col).Value = -19.1
        Case 9
        Cells(7, Col).Value = -16.1
        Case 10
        Cells(7, Col).Value = -13.4
        Case 11
        Cells(7, Col).Value = -10.9
        Case 12
        Cells(7, Col).Value = -8.6
        Case 13
        Cells(7, Col).Value = -6.6
        Case 14
        Cells(7, Col).Value = -4.8
        Case 15
        Cells(7, Col).Value = -3.2
        Case 16
        Cells(7, Col).Value = -1.9
        Case 17
        Cells(7, Col).Value = -0.8
        Case 18
        Cells(7, Col).Value = 0#
        Case 19
        Cells(7, Col).Value = 0.6
        Case 20
        Cells(7, Col).Value = 1#
        Case 21
        Cells(7, Col).Value = 1.2
        Case 22
        Cells(7, Col).Value = 1.3
        Case 23
        Cells(7, Col).Value = 1.2
        Case 24
        Cells(7, Col).Value = 1#
        Case 25
        Cells(7, Col).Value = 0.5
        End Select
    Next Col
    ElseIf SheetType = "TOA" Then
    For Col = 2 To 25
        Select Case Col
        Case 2
        Cells(7, Col).Value = "A Weighting"
        Case 5
        Cells(7, Col).Value = 30.2
        Case 6
        Cells(7, Col).Value = 26.2
        Case 7
        Cells(7, Col).Value = 22.5
        Case 8
        Cells(7, Col).Value = 19.1
        Case 9
        Cells(7, Col).Value = 16.1
        Case 10
        Cells(7, Col).Value = 13.4
        Case 11
        Cells(7, Col).Value = 10.9
        Case 12
        Cells(7, Col).Value = 8.6
        Case 13
        Cells(7, Col).Value = 6.6
        Case 14
        Cells(7, Col).Value = 4.8
        Case 15
        Cells(7, Col).Value = 3.2
        Case 16
        Cells(7, Col).Value = 1.9
        Case 17
        Cells(7, Col).Value = 0.8
        Case 18
        Cells(7, Col).Value = 0#
        Case 19
        Cells(7, Col).Value = -0.6
        Case 20
        Cells(7, Col).Value = -1#
        Case 21
        Cells(7, Col).Value = -1.2
        Case 22
        Cells(7, Col).Value = -1.3
        Case 23
        Cells(7, Col).Value = -1.2
        Case 24
        Cells(7, Col).Value = -1#
        Case 25
        Cells(7, Col).Value = -0.5
        End Select
    Next Col
Else
SheetTypeUnknownError
End If

End Sub

Sub MoveUp(SheetType As String)

Dim StartRw As Integer
Dim endrw As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

StartRw = Selection.Row
endrw = Selection.Row + Selection.Rows.Count - 1

Range("B" & StartRw & ":D" & endrw).Cut Destination:=Range("B" & StartRw - 1 & ":D" & endrw - 1) 'Description

    If Left(SheetType, 3) = "OCT" Then
    Range("E" & StartRw & ":O" & endrw).Cut Destination:=Range("E" & StartRw - 1 & ":O" & endrw - 1) 'Formulas
    Range("B" & StartRw - 1 & ":O" & StartRw - 1).Copy 'formats
    Range("B" & StartRw & ":O" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf Left(SheetType, 2) = "TO" Then
    Range("E" & StartRw & ":AA" & endrw).Cut Destination:=Range("E" & StartRw - 1 & ":AA" & endrw - 1) 'Formulas
    Range("B" & StartRw - 1 & ":AA" & StartRw - 1).Copy 'formats
    Range("B" & StartRw & ":AA" & StartRw).PasteSpecial Paste:=xlPasteFormats
    End If
    
'move to select lower row
Cells(endrw - 1, 2).Select

Application.CutCopyMode = False

End Sub


Sub MoveDown(SheetType As String)

Dim StartRw As Integer
Dim endrw As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

StartRw = Selection.Row
endrw = Selection.Row + Selection.Rows.Count - 1

Range("B" & StartRw & ":D" & endrw).Cut Destination:=Range("B" & StartRw + 1 & ":D" & endrw + 1) 'Description

    If Left(SheetType, 3) = "OCT" Then
    Range("E" & StartRw & ":O" & endrw).Cut Destination:=Range("E" & StartRw + 1 & ":O" & endrw + 1) 'Formulas
    Range("B" & StartRw + 1 & ":O" & StartRw + 1).Copy 'formats
    Range("B" & StartRw & ":O" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf Left(SheetType, 2) = "TO" Then
    Range("E" & StartRw & ":AA" & endrw).Cut Destination:=Range("E" & StartRw + 1 & ":AA" & endrw + 1) 'Formulas
    Range("B" & StartRw + 1 & ":AA" & StartRw + 1).Copy 'formats
    Range("B" & StartRw & ":AA" & StartRw).PasteSpecial Paste:=xlPasteFormats
    End If

'move to select lower row
Cells(endrw + 1, 2).Select

Application.CutCopyMode = False

End Sub

Sub RowReference(SheetType As String)
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmRowReference.Show

    If btnOkPressed = False Then
    End
    End If

SplitAddr = Split(UserSelectedAddress, "$", Len(UserSelectedAddress), vbTextCompare)
    
sheetName = SplitAddr(LBound(SplitAddr)) 'sheet is the first element
rw = CInt(SplitAddr(UBound(SplitAddr))) 'row is the last element

Call ParameterMerge(Selection.Row, SheetType)

Cells(Selection.Row, 2).Value = "Reference to: " & Cells(rw, 2).Value
Cells(Selection.Row, 5).Value = "=" & sheetName & "E$" & rw
ExtendFunction (SheetType)
FormatAs_CellReference (SheetType)
End Sub

Sub SingleCorrection(SheetType As String)
Dim Col As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS
Cells(Selection.Row, 2).Value = "Single Correction"
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=$N" & Selection.Row
    Col = 14
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 5).Value = "=$Z" & Selection.Row
    Col = 26
    End If
Cells(Selection.Row, Col).Value = -5
Call ParameterMerge(Selection.Row, SheetType)
'Cells(Selection.Row, Col).NumberFormat = "0 ""dBA"""
Cells(Selection.Row, Col).NumberFormat = """+""#"" dB"";""-""# ""dB"""
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)
End Sub

Sub AutoSum(SheetType As String)
Dim FindRw As Integer
Dim ScanCol As Integer
Dim FoundRw As Boolean

CheckRow (Selection.Row)
Cells(Selection.Row, 2).Value = "TOTAL SPL"

'find end of range
FindRw = Selection.Row - 1 'one above findrw
ScanCol = Selection.Column
foudnRw = False
    While FoundRw = False
    FindRw = FindRw - 1
    
        If FindRw < 8 Then 'A weighting is on line 7
        'msg = MsgBox("AutoSum Error", vbOKOnly, "ERROR")
        FindRw = 7 'A weighting line is the same as a blank line
        FoundRw = True
        End If
        
        If Cells(FindRw, ScanCol).Value = "" Then FoundRw = True
        
    Wend
    
    If Left(SheetType, 3) = "OCT" Then
    Cells(Selection.Row, 5).Value = "=SUM(E" & FindRw + 1 & ":E" & Selection.Row - 1 & ")"
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 5).Value = "=SUM(E" & FindRw + 1 & ":E" & Selection.Row - 1 & ")" 'same formula!
    Else
    SheetTypeUnknownError
    End If
    
ExtendFunction (SheetType)
    
End Sub

Sub Manual_ExtendFunction(SheetType As String)

CheckRow (Selection.Row)
    
ExtendFunction (SheetType)
    
End Sub

Sub TenLogN(SheetType As String)
CheckRow (Selection.Row)
Cells(Selection.Row, 2).Value = "Multiple sources: 10log(n)"

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=10*LOG(" & Cells(Selection.Row, 14).Address(False, True) & ")"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 14) = 2
    Cells(Selection.Row, 14).NumberFormat = """n = ""0"
    
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    Cells(Selection.Row, 5).Value = "=10*LOG(" & Cells(Selection.Row, 26).Address(False, True) & ")"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 26) = 2
    Cells(Selection.Row, 26).NumberFormat = """n = ""0"
    
    Else
    SheetTypeUnknownError
    End If
    
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)
End Sub

Sub TenLogOneOnT(SheetType As String)
CheckRow (Selection.Row)
Cells(Selection.Row, 2).Value = "Time Correction: 10log(1/t)"

    If Left(SheetType, 3) = "OCT" Then 'oct or OCT
    Cells(Selection.Row, 5).Value = "=10*LOG(1/" & Cells(Selection.Row, 14).Address(False, True) & ")"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 14) = 2
    Cells(Selection.Row, 14).NumberFormat = """n = 1/""0"
    
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    Cells(Selection.Row, 5).Value = "=10*LOG(1/" & Cells(Selection.Row, 26).Address(False, True) & ")"
    Call ParameterMerge(Selection.Row, SheetType)
    Cells(Selection.Row, 26) = 2
    Cells(Selection.Row, 26).NumberFormat = """n = 1/""0"
    
    Else
    SheetTypeUnknownError
    End If
    
ExtendFunction (SheetType)
UserInputFormat_ParamCol (SheetType)
End Sub

Sub OneThirdsToOctave(SheetType As String)

Dim SplitAddr() As String
Dim sheetName As String
Dim rw As Integer
Dim refCol As Integer
Dim targetRange As String

CheckRow (Selection.Row)

    If Left(SheetType, 3) = "OCT" Then 'oct or OCTA
    
    frmClickAddress.Show

        If btnOkPressed = False Then
        End
        End If
        
    Cells(Selection.Row, 2).Value = "Import from TO"
    
    SplitAddr = Split(UserSelectedAddress, "$", Len(UserSelectedAddress), vbTextCompare)
    
    sheetName = SplitAddr(LBound(SplitAddr)) 'sheet is the first element
    rw = CInt(SplitAddr(UBound(SplitAddr))) 'row is the last element
    
        refCol = 5
        For Col = 6 To 12
        targetRange = Range(Cells(rw, refCol), Cells(rw, refCol + 2)).Address(False, False)
            If SumOrAverage = "Sum" Then
            Cells(Selection.Row, Col).Value = "=SPLSUM(" & sheetName & targetRange & ")"
            ElseIf SumOrAverage = "Average" Then
            Cells(Selection.Row, Col).Value = "=AVERAGE(" & sheetName & targetRange & ")"
            Else
            msg = MsgBox("What did you press? Option does not exist.", vbOKOnly, "ERROR!")
            End If
        refCol = refCol + 3
        Next Col
        
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    msg = MsgBox("Imports into OCT or OCTA", vbOKOnly, "WRONG WAY GO BACK!")
    Else
    SheetTypeUnknownError
    End If

End Sub


Sub ConvertToAWeight(SheetType As String)
CheckRow (Selection.Row)

Cells(Selection.Row, 5).Value = "=E" & Selection.Row - 1 & "+E$7"

ExtendFunction (SheetType)

    If Right(SheetType, 1) = "A" Then
    Cells(Selection.Row, 2).Value = "Linear Spectrum"
    Else
    Cells(Selection.Row, 2).Value = "A Weighted Spectrum"
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Called from other functions, no check needed from this point onwards
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ExtendFunction(SheetType As String)
Dim rw As Integer
Dim StartAddr As String
rw = Selection.Row
StartAddr = Selection.Address
    If Left(SheetType, 3) = "OCT" Then 'OCT or OCTA
    Cells(rw, 5).Copy
    Range(Cells(rw, 5), Cells(rw, 13)).PasteSpecial Paste:=xlPasteFormulas
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    Cells(rw, 5).Copy
    Range(Cells(rw, 5), Cells(rw, 25)).PasteSpecial Paste:=xlPasteFormulas
    Else
    SheetTypeUnknownError
    End
    End If
Application.CutCopyMode = False
Range(StartAddr).Select
End Sub

Sub ParameterMerge(rw As Integer, SheetType As String)
    If Left(SheetType, 3) = "OCT" Then 'OCT or OCTA
        If Cells(rw, 14).MergeCells = False Then
        Range(Cells(rw, 14), Cells(rw, 15)).Merge
        Cells(rw, 14).HorizontalAlignment = xlCenter
        Cells(rw, 14).VerticalAlignment = xlCenter
        End If
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
        If Cells(rw, 26).MergeCells = False Then
        Range(Cells(rw, 26), Cells(rw, 27)).Merge
        Cells(rw, 26).HorizontalAlignment = xlCenter
        Cells(rw, 26).VerticalAlignment = xlCenter
        End If
    Else
        SheetTypeUnknownError
        End
    End If
End Sub

Sub ParameterUnmerge(rw As Integer, SheetType As String)
    If Left(SheetType, 3) = "OCT" Then 'OCT or OCTA
        If Cells(rw, 14).MergeCells Then 'cells are merged, unmerge
        Range(Cells(rw, 14), Cells(rw, 15)).UnMerge
        Range(Cells(rw, 14), Cells(rw, 15)).Borders.LineStyle = xlContinuous
        End If
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
        If Cells(rw, 26).MergeCells Then 'cells are merged, unmerge
        Range(Cells(rw, 26), Cells(rw, 27)).UnMerge
        Range(Cells(rw, 26), Cells(rw, 27)).Borders.LineStyle = xlContinuous
        End If
    Else
        SheetTypeUnknownError
        End
    End If
End Sub

Sub UserInputFormat(SheetType As String)
    If Left(SheetType, 3) = "OCT" Then
    Range(Cells(Selection.Row, 5), Cells(Selection.Row, 13)).Interior.Color = RGB(251, 251, 143)
    ElseIf Left(SheetType, 2) = "TO" Then
    Range(Cells(Selection.Row, 5), Cells(Selection.Row, 25)).Interior.Color = RGB(251, 251, 143)
    Else
    End If
End Sub

Sub UserInputFormat_ParamCol(SheetType As String)
    If Left(SheetType, 3) = "OCT" Then
    Range(Cells(Selection.Row, 14), Cells(Selection.Row, 15)).Interior.Color = RGB(251, 251, 143)
    ElseIf Left(SheetType, 2) = "TO" Then
    Range(Cells(Selection.Row, 26), Cells(Selection.Row, 27)).Interior.Color = RGB(251, 251, 143)
    Else
    SheetTypeUnknownError
    End If
End Sub

Sub SheetTypeUnknownError()
msg = MsgBox("Sheet Type Unknown", vbOKOnly, "ERROR")
End Sub
