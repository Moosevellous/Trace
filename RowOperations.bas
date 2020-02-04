Attribute VB_Name = "RowOperations"
Public UserSelectedAddress As String
Public SumAverageMode As String
Public LookupMultiRow As Boolean

Public Sub CheckRow(rw As Integer)
'Checks that user isn't in header rows. These rows are protected by this function. None shall Pass.
If rw <= 7 Then End
End Sub

Sub ClearRw(SheetType As String)
Dim rw As Integer
Dim bandsStart As Integer
Dim bandsEnd As Integer
Dim parameterCol1 As Integer
Dim parameterCol2 As Integer
Dim commentCol As Integer
Dim hasTypeCode As Boolean
Dim hasParamCol As Boolean


CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Selection.Rows.Count > 1 Then
    msg = MsgBox("Are you sure you want to clear rows?", vbYesNo, "Check")
    Else
    msg = vbYes
    End If

TypeCode = False

    If msg = vbYes Then
        For rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1

            If Left(SheetType, 3) = "OCT" Then
            
            hasTypeCode = True
            bandsStart = 5
            bandsEnd = 13
            hasParamCol = True
            ParamCol1 = 14
            ParamCol2 = 15
            
            ElseIf Left(SheetType, 2) = "TO" Then
            
            hasTypeCode = True
            bandsStart = 5
            bandsEnd = 25
            hasParamCol = True
            ParamCol1 = 26
            ParamCol2 = 27
            
         
            ElseIf SheetType = "LF_TO" Then
            
            hasTypeCode = True
            bandsStart = 5
            bandsEnd = 31
            hasParamCol = True
            ParamCol1 = 32
            ParamCol2 = 33
            
            
            ElseIf SheetType = "CVT" Then
            
            hasTypeCode = True
            bandsStart = 5
            bandsEnd = 44
            hasParamCol = False
            
            End If
            
            'apply style
            If hasTypeCode Then
            'description/comment
            Cells(rw, 2).ClearContents
            Cells(rw, 2).ClearComments
            Cells(rw, 2).Validation.Delete 'for dropdown boxes
                
                'PARAMETER COLUMNS
                If hasParamCol = True Then
                'values
                Range(Cells(rw, bandsStart), Cells(rw, ParamCol2)).ClearContents
                Range(Cells(rw, bandsStart), Cells(rw, ParamCol2)).Font.colorindex = 0
                Range(Cells(rw, bandsStart), Cells(rw, ParamCol2)).Interior.colorindex = 0 'no colour
                'parameter columns
                Range(Cells(rw, ParamCol1), Cells(rw, ParamCol2)).UnMerge
                Cells(rw, ParamCol1).Validation.Delete 'for dropdown boxes
                Cells(rw, ParamCol2).Validation.Delete 'for dropdown boxes
                Cells(rw, ParamCol1).ClearComments
                Cells(rw, ParamCol2).ClearComments
                Cells(rw, ParamCol1).NumberFormat = "General"
                Cells(rw, ParamCol2).NumberFormat = "General"
                Else
                Range(Cells(rw, bandsStart), Cells(rw, bandsEnd)).ClearContents
                Range(Cells(rw, bandsStart), Cells(rw, bandsEnd)).Font.colorindex = 0
                Range(Cells(rw, bandsStart), Cells(rw, bandsEnd)).Interior.colorindex = 0 'no colour
                End If
                
            'Debug.Print Cells(rw, bandsStart).NumberFormat
                'detect scientific notation
                If InStr(1, Cells(rw, bandsStart).NumberFormat, "E", vbTextCompare) > 0 Then
                Range(Cells(rw, bandsStart), Cells(rw, bandsEnd)).NumberFormat = "0.0"
                End If
            Range(Cells(rw, 2), Cells(rw, bandsEnd)).FormatConditions.Delete 'removes heatmap
            ApplyTraceStyle "Trace Normal", SheetType, rw
            'standard formatting, column 2 is bold
            Cells(rw, 4).Font.Bold = True
            Else
            SheetTypeUnknownError (SheetType)
            End If

        Next rw
    End If

End Sub

Sub FlipSign(SheetType As String, Optional SkipUserInput As Boolean)
Dim rw As Integer
Dim StartCol As Integer
Dim EndCol As Integer
CheckRow (Selection.Row)

    If SkipUserInput = False Then
    ApplyToRow = MsgBox("Apply to current row? (Selecting 'No' will apply to selection only)", vbYesNoCancel, "Flip Sign")
    Else 'skip, default to entire row
    ApplyToRow = vbYes
    End If
    

    If ApplyToRow = vbYes Then
    StartCol = GetSheetTypeRanges(SheetType, "DataStart")
    EndCol = GetSheetTypeRanges(SheetType, "DataEnd")
    ElseIf ApplyToRow = vbNo Then
    StartCol = Selection.Column
    EndCol = Selection.Column + Selection.Columns.Count - 1
    Else 'cancel or anything else
    End
    End If

    For rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
        For Col = StartCol To EndCol
            If Cells(rw, Col).HasFormula Then
                If Mid(Cells(rw, Col).Formula, 2, 1) = "-" Then 'second character of formula
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

Sub A_weight_Corrections(SheetType As String)

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
SheetTypeUnknownError (SheetType)
End If

End Sub

Sub MoveUp(SheetType As String)

Dim StartRw As Integer
Dim EndRw As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

Application.ScreenUpdating = False

StartRw = Selection.Row
EndRw = Selection.Row + Selection.Rows.Count - 1

Range("B" & StartRw & ":D" & EndRw).Cut Destination:=Range("B" & StartRw - 1 & ":D" & EndRw - 1) 'Description

    If Left(SheetType, 3) = "OCT" Then
    Range("E" & StartRw & ":O" & EndRw).Cut Destination:=Range("E" & StartRw - 1 & ":O" & EndRw - 1) 'Formulas
    Range("B" & StartRw - 1 & ":O" & StartRw - 1).Copy 'formats
    Range("B" & StartRw & ":O" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf Left(SheetType, 2) = "TO" Then
    Range("E" & StartRw & ":AA" & EndRw).Cut Destination:=Range("E" & StartRw - 1 & ":AA" & EndRw - 1) 'Formulas
    Range("B" & StartRw - 1 & ":AA" & StartRw - 1).Copy 'formats
    Range("B" & StartRw & ":AA" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf SheetType = "LF_TO" Then
    Range("E" & StartRw & ":AG" & EndRw).Cut Destination:=Range("E" & StartRw - 1 & ":AG" & EndRw - 1) 'Formulas
    Range("B" & StartRw - 1 & ":AG" & StartRw - 1).Copy 'formats
    Range("B" & EndRw & ":AG" & EndRw).PasteSpecial Paste:=xlPasteFormats
    End If
    
'move to select lower row
Range(Cells(StartRw - 1, 2), Cells(EndRw - 1, 2)).Select

Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub


Sub MoveDown(SheetType As String)

Dim StartRw As Integer
Dim EndRw As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

Application.ScreenUpdating = False

StartRw = Selection.Row
EndRw = Selection.Row + Selection.Rows.Count - 1

Range("B" & StartRw & ":D" & EndRw).Cut Destination:=Range("B" & StartRw + 1 & ":D" & EndRw + 1) 'Description

    If Left(SheetType, 3) = "OCT" Then
    Range("E" & StartRw & ":O" & EndRw).Cut Destination:=Range("E" & StartRw + 1 & ":O" & EndRw + 1) 'Formulas
    Range("B" & StartRw + 1 & ":O" & StartRw + 1).Copy 'formats
    Range("B" & StartRw & ":O" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf Left(SheetType, 2) = "TO" Then
    Range("E" & StartRw & ":AA" & EndRw).Cut Destination:=Range("E" & StartRw + 1 & ":AA" & EndRw + 1) 'Formulas
    Range("B" & StartRw + 1 & ":AA" & StartRw + 1).Copy 'formats
    Range("B" & StartRw & ":AA" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf SheetType = "LF_TO" Then
    Range("E" & StartRw & ":AG" & EndRw).Cut Destination:=Range("E" & StartRw + 1 & ":AG" & EndRw + 1) 'Formulas
    Range("B" & StartRw + 1 & ":AG" & StartRw + 1).Copy 'formats
    Range("B" & StartRw & ":AG" & StartRw).PasteSpecial Paste:=xlPasteFormats
    End If

'move to select lower row
Range(Cells(StartRw + 1, 2), Cells(EndRw + 1, 2)).Select

Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub

Sub RowReference(SheetType As String)
Dim FirstRow As Integer
Dim LastRow As Integer
Dim SheetName As String

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmRowReference.Show

    If btnOkPressed = False Then
    End
    End If

    If UserSelectedAddress = "" Then End 'error catch
    
    SheetName = GetSheetName(UserSelectedAddress)
    FirstRow = GetFirstRow(UserSelectedAddress)
    LastRow = GetLastRow(UserSelectedAddress)
    
    If LookupMultiRow = False Then
    Cells(Selection.Row, 2).Value = "=CONCAT(""Ref: ""," & SheetName & "$B$" & FirstRow & ")"
    Cells(Selection.Row, 5).Value = "=" & SheetName & "E$" & FirstRow
    ExtendFunction (SheetType)
    
    Else 'multimode = true
    
    'data validation
    With Cells(Selection.Row, 2).Validation
         .Delete
         .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
         xlBetween, Formula1:="=" & SheetName & "$B$" & FirstRow & ":$B$" & LastRow
         .IgnoreBlank = True
         .InCellDropdown = True
         .InputTitle = ""
         .ErrorTitle = ""
         .InputMessage = ""
         .ErrorMessage = ""
         .ShowInput = True
         .ShowError = True
     End With
    
    'select first entry by default
    Cells(Selection.Row, 2).Value = Range(SheetName & "$B$" & FirstRow)
    
    'create index-match formula
    If Left(SheetType, 3) = "OCT" Then
    Debug.Print "=INDEX(" & SheetName & "$E$" & FirstRow & ":$M$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
    "," & SheetName & "$B$" & FirstRow & ":$B$" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!E$6," & SheetName & "$E$6:$M$6,0))"
    Cells(Selection.Row, 5).Value = "=INDEX(" & SheetName & "$E$" & FirstRow & ":$M$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
    "," & SheetName & "$B$" & FirstRow & ":$B$" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!E$6," & SheetName & "$E$6:$M$6,0))" '<----note that SheetName includes apostrophe character and ActiveSheet.Name does not.....trickyyyyy
    ElseIf Left(SheetType, 2) = "TO" Then
    Cells(Selection.Row, 5).Value = "=INDEX(" & SheetName & "!$E$" & FirstRow & ":$Y$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'$B$" & Selection.Row & _
    "," & SheetName & "$B$" & FirstRow & ":$B$" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'E$6," & SheetName & "$E$6:$Y$6,0))"
    End If
    
    ExtendFunction (SheetType)
    
    End If
    
    
'apply Trace Reference style
fmtReference (SheetType)  'OLD VERSION: FormatAs_CellReference (SheetType)

End Sub

Function GetSheetName(inputStr As String) 'Sheet name, first row, last row
Dim splitStr() As String
splitStr = Split(inputStr, "!", Len(inputStr), vbTextCompare)
    If Right(splitStr(0), 1) = "!" Then
    GetSheetName = splitStr(0)
    Else
    GetSheetName = splitStr(0) & "!" 'sheet is the first element
    End If
End Function

Function GetFirstRow(inputStr As String)
Dim splitStr() As String
splitStr = Split(inputStr, "$", Len(inputStr), vbTextCompare)
    If Right(splitStr(2), 1) = ":" Then
    GetFirstRow = CInt(Left(splitStr(2), Len(splitStr(2)) - 1)) 'trim one colon character = colonoscopy???
    Else
    GetFirstRow = CInt(splitStr(2))
    End If
End Function

Function GetLastRow(inputStr As String)
Dim splitStr() As String
splitStr = Split(inputStr, "$", Len(inputStr), vbTextCompare)
GetLastRow = CInt(splitStr(UBound(splitStr)))
End Function

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
fmtUserInput SheetType, True
End Sub

Sub AutoSum(SheetType As String)
Dim FindRw As Integer
Dim ScanCol As Integer
Dim FoundRw As Boolean

CheckRow (Selection.Row)
Cells(Selection.Row, 2).Value = "Total"

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
    SheetTypeUnknownError (SheetType)
    End If
    
ExtendFunction (SheetType)

fmtTotal (SheetType)
    
End Sub

Sub Manual_ExtendFunction(SheetType As String)

CheckRow (Selection.Row)
    
ExtendFunction (SheetType)
    
End Sub


Sub OneThirdsToOctave(SheetType As String)

Dim splitAddr() As String
Dim SheetName As String
Dim rw As Integer
Dim refCol As Integer
Dim targetRange As String

CheckRow (Selection.Row)

    If Left(SheetType, 3) = "OCT" Then 'oct or OCTA
    
    frmConvert.Show

        If btnOkPressed = False Then
        End
        End If
        
    Cells(Selection.Row, 2).Value = "Import from TO"
    
    If UserSelectedAddress = "" Then End
    
    splitAddr = Split(UserSelectedAddress, "$", Len(UserSelectedAddress), vbTextCompare)
    
    SheetName = splitAddr(LBound(splitAddr)) 'sheet is the first element
    rw = CInt(splitAddr(UBound(splitAddr))) 'row is the last element
    
        refCol = 5
        For Col = 6 To 12
        targetRange = Range(Cells(rw, refCol), Cells(rw, refCol + 2)).Address(False, False)
            Select Case SumAverageMode 'selected from radio boxes in form frmConvert
            Case Is = "Sum"
            Cells(Selection.Row, Col).Value = "=SPLSUM(" & SheetName & targetRange & ")"
            Case Is = "Average"
            Cells(Selection.Row, Col).Value = "=AVERAGE(" & SheetName & targetRange & ")"
            Case Is = "Log Av"
            Cells(Selection.Row, Col).Value = "=SPLAV(" & SheetName & targetRange & ")"
            Case Is = "TL" 'positive spectra returned as positive
            Cells(Selection.Row, Col).Value = "=TL_ThirdsToOctave(" & SheetName & targetRange & ")"
            '"=-10*LOG((1/3)*(10^(-" & sheetName & Cells(rw, refCol).Address & "/10)+10^(-" & sheetName & Cells(rw, refCol + 1).Address & "/10)+10^(-" & sheetName & Cells(rw, refCol + 2).Address & "/10)))"
            End Select
        refCol = refCol + 3
        Next Col
        
    'apply reference style
    fmtReference (SheetType)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf SheetType = "CVT" Then '<------- CONVERSION SHEET TYPE
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    frmConvert.refRangeSelector.Enabled = False
    frmConvert.refRangeSelector.Value = Selection.Address  'OLD VERSION: "'" & ActiveSheet.Name & "'!" & Selection.Address
    frmConvert.Show

        If btnOkPressed = False Then
        End
        End If
    
    If UserSelectedAddress = "" Then End
    
    splitAddr = Split(UserSelectedAddress, "$", Len(UserSelectedAddress), vbTextCompare)
    
    'SheetName = splitAddr(LBound(splitAddr)) 'sheet is the first element
    rw = CInt(splitAddr(UBound(splitAddr))) 'row is the last element, eg $A$1
    
        refCol = 5
        For Col = 33 To 41
        targetRange = Range(Cells(rw, refCol), Cells(rw, refCol + 2)).Address(False, False)
            Select Case SumAverageMode 'selected from radio boxes in form frmConvert
            Case Is = "Sum"
            Cells(Selection.Row, Col).Value = "=SPLSUM(" & SheetName & targetRange & ")"
            Case Is = "Average"
            Cells(Selection.Row, Col).Value = "=AVERAGE(" & SheetName & targetRange & ")"
            Case Is = "Log Av"
            Cells(Selection.Row, Col).Value = "=SPLAV(" & SheetName & targetRange & ")"
            Case Is = "TL" 'positive spectra returned as positive
            Cells(Selection.Row, Col).Value = "=TL_ThirdsToOctave(" & SheetName & targetRange & ")"
            End Select
        refCol = refCol + 3
        Next Col
        
    'apply reference style ''''''''''''''maybe dont for CVT sheet?
    'fmtReference (SheetType)
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
    msg = MsgBox("Imports into OCT or OCTA", vbOKOnly, "WRONG WAY GO BACK!")
    Else
    SheetTypeUnknownError (SheetType)
    End If

End Sub


Sub ConvertToAWeight(SheetType As String)
CheckRow (Selection.Row)

Cells(Selection.Row, 5).Value = "=E" & Selection.Row - 1 & "+E$7"

ExtendFunction (SheetType)
    If Left(SheetType, 2) = "OCT" Or Left(SheetType, 2) = "TO" Then
        If Right(SheetType, 1) = "A" Then
        Cells(Selection.Row, 2).Value = "Linear Spectrum"
        Else
        Cells(Selection.Row, 2).Value = "A Weighted Spectrum"
        End If
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
    ElseIf SheetType = "LF_TO" Then
    Cells(rw, 5).Copy
    Range(Cells(rw, 5), Cells(rw, 31)).PasteSpecial Paste:=xlPasteFormulas
    Else
    SheetTypeUnknownError (SheetType)
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
    ElseIf SheetType = "LF_TO" Then
        If Cells(rw, 32).MergeCells = False Then
        Range(Cells(rw, 32), Cells(rw, 33)).Merge
        Cells(rw, 32).HorizontalAlignment = xlCenter
        Cells(rw, 32).VerticalAlignment = xlCenter
        End If
    Else
        SheetTypeUnknownError (SheetType)
        End
    End If
End Sub

Sub ParameterUnmerge(rw As Integer, SheetType As String)
    If Left(SheetType, 3) = "OCT" Then 'OCT or OCTA
        If Cells(rw, 14).MergeCells Then 'cells are merged, unmerge
        Range(Cells(rw, 14), Cells(rw, 15)).UnMerge
        Range(Cells(rw, 14), Cells(rw, 15)).Borders.LineStyle = xlContinuous
        Selection.Borders(xlInsideVertical).Weight = xlHairline
        End If
    ElseIf Left(SheetType, 2) = "TO" Then 'TO or TOA
        If Cells(rw, 26).MergeCells Then 'cells are merged, unmerge
        Range(Cells(rw, 26), Cells(rw, 27)).UnMerge
        Range(Cells(rw, 26), Cells(rw, 27)).Borders.LineStyle = xlContinuous
        End If
    ElseIf SheetType = "LF_TO" Then
        If Cells(rw, 32).MergeCells Then 'cells are merged, unmerge
        Range(Cells(rw, 32), Cells(rw, 33)).UnMerge
        Range(Cells(rw, 32), Cells(rw, 33)).Borders.LineStyle = xlContinuous
        End If
    Else
        SheetTypeUnknownError (SheetType)
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

Sub UserInputFormat_ParamCol(SheetType As String) 'legacy code, will redirect to new function

fmtUserInput (SheetType)
    
'OLD VERSION, USES COLOURS, NOT STYLES
'    If Left(SheetType, 3) = "OCT" Then
'    Range(Cells(Selection.Row, 14), Cells(Selection.Row, 15)).Interior.Color = RGB(251, 251, 143)
'    ElseIf Left(SheetType, 2) = "TO" Then
'    Range(Cells(Selection.Row, 26), Cells(Selection.Row, 27)).Interior.Color = RGB(251, 251, 143)
'    ElseIf SheetType = "LF_TO" Then
'    Range(Cells(Selection.Row, 32), Cells(Selection.Row, 33)).Interior.Color = RGB(251, 251, 143)
'    Else
'    SheetTypeUnknownError(SheetType)
'    End If

End Sub

Sub SheetTypeUnknownError(SheetType As String)
msg = MsgBox("Not implemented for Typecode: " & SheetType, vbOKOnly, "Error - Sheet Type")
End
End Sub
