Attribute VB_Name = "RowOperations"
Public UserSelectedAddress As String
Public SumAverageMode As String

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


CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Selection.Rows.count > 1 Then
    msg = MsgBox("Are you sure you want to clear rows?", vbYesNo, "Check")
    Else
    msg = vbYes
    End If

TypeCode = False

    If msg = vbYes Then
        For rw = Selection.Row To Selection.Row + Selection.Rows.count - 1

            If Left(SheetType, 3) = "OCT" Then
            
            hasTypeCode = True
            bandsStart = 5
            bandsEnd = 13
            ParamCol1 = 14
            ParamCol2 = 15
            
            ElseIf Left(SheetType, 2) = "TO" Then
            
            hasTypeCode = True
            bandsStart = 5
            bandsEnd = 25
            ParamCol1 = 26
            ParamCol2 = 27
            
         
            ElseIf SheetType = "LF_TO" Then
            
            hasTypeCode = True
            bandsStart = 5
            bandsEnd = 31
            ParamCol1 = 32
            ParamCol2 = 33
            
            End If
            
            'apply style
            If hasTypeCode Then
            'description/comment
            Cells(rw, 2).ClearContents
            Cells(rw, 2).ClearComments
            'values
            Range(Cells(rw, bandsStart), Cells(rw, ParamCol2)).ClearContents
            Range(Cells(rw, bandsStart), Cells(rw, ParamCol2)).Font.ColorIndex = 0
            Range(Cells(rw, bandsStart), Cells(rw, ParamCol2)).Interior.ColorIndex = 0 'no colour
            Range(Cells(rw, ParamCol1), Cells(rw, ParamCol2)).UnMerge
            Cells(rw, ParamCol1).Validation.Delete 'for dropdown boxes
            Cells(rw, ParamCol2).Validation.Delete 'for dropdown boxes
            Cells(rw, ParamCol1).ClearComments
            Cells(rw, ParamCol2).ClearComments
            Cells(rw, ParamCol1).NumberFormat = "General"
            Cells(rw, ParamCol2).NumberFormat = "General"
            Range(Cells(rw, 2), Cells(rw, bandsEnd)).FormatConditions.Delete 'removes heatmap
            ApplyTraceStyle "Trace Normal", SheetType, rw
            'standard formatting, column 2 is bold
            Cells(rw, 4).Font.Bold = True
            Else
            msg = MsgBox("Not implemented for this Typecode: " & TypeCode, vbOKOnly, "Error - Sheet Type")
            End If

        Next rw
    End If

End Sub

Sub FlipSign(SheetType As String)
Dim rw As Integer
CheckRow (Selection.Row)

    For rw = Selection.Row To Selection.Row + Selection.Rows.count - 1
        For Col = Selection.Column To Selection.Column + Selection.Columns.count - 1
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
SheetTypeUnknownError
End If

End Sub

Sub MoveUp(SheetType As String)

Dim StartRw As Integer
Dim EndRw As Integer

CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

Application.ScreenUpdating = False

StartRw = Selection.Row
EndRw = Selection.Row + Selection.Rows.count - 1

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
EndRw = Selection.Row + Selection.Rows.count - 1

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
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

frmRowReference.Show

    If btnOkPressed = False Then
    End
    End If

SplitAddr = Split(UserSelectedAddress, "$", Len(UserSelectedAddress), vbTextCompare)

If UserSelectedAddress = "" Then End 'error catch

sheetName = SplitAddr(LBound(SplitAddr)) 'sheet is the first element
'sheetNameShort = Mid(sheetName, 2, Len(sheetName) - 3) 'trim extra characters in the string
'sheetnameshort = Left(sheetName, Len(sheetName) - 1)

rw = CInt(SplitAddr(UBound(SplitAddr))) 'row is the last element

Call ParameterMerge(Selection.Row, SheetType)

Cells(Selection.Row, 2).Value = "=CONCAT(""Reference to: ""," & sheetName & "$B$" & rw & ")"
Cells(Selection.Row, 5).Value = "=" & sheetName & "E$" & rw
ExtendFunction (SheetType)
'apply Trace Reference style
fmtReference (SheetType)  'OLD VERSION: FormatAs_CellReference (SheetType)
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
fmtUserInput SheetType, True
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

fmtTotal (SheetType)
    
End Sub

Sub Manual_ExtendFunction(SheetType As String)

CheckRow (Selection.Row)
    
ExtendFunction (SheetType)
    
End Sub


Sub OneThirdsToOctave(SheetType As String)

Dim SplitAddr() As String
Dim sheetName As String
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
    
    SplitAddr = Split(UserSelectedAddress, "$", Len(UserSelectedAddress), vbTextCompare)
    
    sheetName = SplitAddr(LBound(SplitAddr)) 'sheet is the first element
    rw = CInt(SplitAddr(UBound(SplitAddr))) 'row is the last element
    
        refCol = 5
        For Col = 6 To 12
        targetRange = Range(Cells(rw, refCol), Cells(rw, refCol + 2)).Address(False, False)
            Select Case SumAverageMode 'selected from radio boxes in form frmConvert
            Case Is = "Sum"
            Cells(Selection.Row, Col).Value = "=SPLSUM(" & sheetName & targetRange & ")"
            Case Is = "Average"
            Cells(Selection.Row, Col).Value = "=AVERAGE(" & sheetName & targetRange & ")"
            Case Is = "Log Av"
            Cells(Selection.Row, Col).Value = "=SPLAV(" & sheetName & targetRange & ")"
            Case Is = "TL" 'positive spectra returned as positive
            Cells(Selection.Row, Col).Value = "=TL_ThirdsToOctave(" & sheetName & targetRange & ")"
            '"=-10*LOG((1/3)*(10^(-" & sheetName & Cells(rw, refCol).Address & "/10)+10^(-" & sheetName & Cells(rw, refCol + 1).Address & "/10)+10^(-" & sheetName & Cells(rw, refCol + 2).Address & "/10)))"
            End Select
        refCol = refCol + 3
        Next Col
        
    'apply reference style
    fmtReference (SheetType)
        
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
    ElseIf SheetType = "LF_TO" Then
    Cells(rw, 5).Copy
    Range(Cells(rw, 5), Cells(rw, 31)).PasteSpecial Paste:=xlPasteFormulas
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
    ElseIf SheetType = "LF_TO" Then
        If Cells(rw, 32).MergeCells = False Then
        Range(Cells(rw, 32), Cells(rw, 33)).Merge
        Cells(rw, 32).HorizontalAlignment = xlCenter
        Cells(rw, 32).VerticalAlignment = xlCenter
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
'    SheetTypeUnknownError
'    End If

End Sub

Sub SheetTypeUnknownError()
msg = MsgBox("Sheet Type Unknown", vbOKOnly, "ERROR")
End Sub
