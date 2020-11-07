Attribute VB_Name = "RowOperations"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
Public UserSelectedAddress As String
Public SumAverageMode As String
Public LookupMultiRow As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'==============================================================================
' Name:     GetSheetName
' Author:   PS
' Desc:     Returns the name of the sheet
' Args:     inputStr - Range string
' Comments: (1)
'==============================================================================
Function GetSheetName(InputStr As String) 'Sheet name, first row, last row
Dim SplitStr() As String
SplitStr = Split(InputStr, "!", Len(InputStr), vbTextCompare)
    If Right(SplitStr(0), 1) = "!" Then
    GetSheetName = SplitStr(0)
    Else
    GetSheetName = SplitStr(0) & "!" 'sheet is the first element
    End If
End Function

'==============================================================================
' Name:     GetFirstRow
' Author:   PS
' Desc:     Returns the first row of the input range
' Args:     inputStr - Range string
' Comments: (1)
'==============================================================================
Function GetFirstRow(InputStr As String)
Dim SplitStr() As String
SplitStr = Split(InputStr, "$", Len(InputStr), vbTextCompare)
    If Right(SplitStr(2), 1) = ":" Then
    'trim one colon character = colonoscopy???
    GetFirstRow = CInt(Left(SplitStr(2), Len(SplitStr(2)) - 1))
    Else
    GetFirstRow = CInt(SplitStr(2))
    End If
End Function

'==============================================================================
' Name:     GetLastRow
' Author:   PS
' Desc:     Returns the last row of the input range
' Args:     inputStr - Range string
' Comments: (1)
'==============================================================================
Function GetLastRow(InputStr As String)
Dim SplitStr() As String
SplitStr = Split(InputStr, "$", Len(InputStr), vbTextCompare)
GetLastRow = CInt(SplitStr(UBound(SplitStr)))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'==============================================================================
' Name:     ClearRow
' Author:   PS
' Desc:     Deletes contents and clears formatting for selected ranges
' Args:     None
' Comments: (1) Also called by other functions
'           (2) SkipUserInput defaults to false
'==============================================================================
Sub ClearRow(Optional SkipUserInput As Boolean)
Dim rw As Integer
Dim P As Integer
Dim i As Integer
    
    If SkipUserInput = False Then 'runs by default
        If Selection.Rows.Count > 1 Then
        ClearMultipleRows = MsgBox("Are you sure you want to clear all rows?", _
            vbYesNo, "Check, one, two...")
        End If
        
        If ClearMultipleRows = vbNo Then End 'catch user abort
    End If

Application.ScreenUpdating = False
    
    'loop for clearing data, works
    For rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
    
    'Description
    Cells(rw, T_Description).ClearContents
    Cells(rw, T_Description).ClearComments
    Cells(rw, T_Description).Validation.Delete 'for dropdown boxes
        
    'LossGain values
    Range(Cells(rw, T_LossGainStart), Cells(rw, T_LossGainEnd)).ClearContents
        
        'Regenerated noise columns
        If T_RegenStart <> -1 Then
        Range(Cells(rw, T_RegenStart), Cells(rw, T_RegenEnd)).ClearContents
        End If
       
        'Comment
        If T_Comment <> -1 Then
        Cells(rw, T_Comment).ClearContents
        End If
    
        'Parameter columns
        If T_ParamStart >= 0 Then
        ParameterUnmerge (rw)
            For P = T_ParamStart To T_ParamEnd
            Cells(rw, P).Validation.Delete
            Cells(rw, P).ClearContents
            Cells(rw, P).ClearComments
            Cells(rw, P).NumberFormat = "General"
            Next P
            
            'sparklines
            For i = 0 To 3
            Cells(rw, T_ParamStart + i).SparklineGroups.Clear
            Next i
        End If
        
    'remove heatmap
    Range(Cells(rw, T_Description), Cells(rw, T_LossGainEnd)).FormatConditions.Delete
    
    'apply styles and reset units
    ApplyTraceStyle "Trace Normal", rw
    SetUnits "Clear", T_LossGainStart, 0, T_LossGainEnd
    'styles for parameter columns
        If T_ParamStart >= 0 Then
        ApplyTraceStyle "Trace Normal", rw, True
        End If
        
    'bold to overall column
    Cells(rw, T_LossGainStart - 1).Font.Bold = True
    
    'lock cells
    Range(Cells(rw, T_Description), Cells(rw, T_LossGainEnd)).Locked = True
    
    Next rw
Application.ScreenUpdating = True

End Sub

'==============================================================================
' Name:     FlipSign
' Author:   PS
' Desc:     Makes selected calls (or row) negative values. Maintains the
'           formula and will revert back.
' Args:     None
' Comments: (1) SkipUserInput defaults to false
'==============================================================================
Sub FlipSign()
Dim rw As Integer
Dim startCol As Integer
Dim endCol As Integer

    For rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
        For col = T_LossGainStart To T_LossGainEnd
            If Cells(rw, col).HasFormula Then
                'check if second character of formula is minus
                If Mid(Cells(rw, col).Formula, 2, 1) = "-" Then
                Cells(rw, col).Formula = Replace(Cells(rw, col).Formula, _
                    "=-", "=", 1, Len(Cells(rw, col).Formula), vbTextCompare)
                Else
                Cells(rw, col).Formula = Replace(Cells(rw, col).Formula, _
                    "=", "=-", 1, Len(Cells(rw, col).Formula), vbTextCompare)
                End If
            Else
                Cells(rw, col).Value = Cells(rw, col).Value * -1
            End If
        Next col
    Next rw

End Sub

'==============================================================================
' Name:     A_weight_Corrections
' Author:   PS
' Desc:     Returns value of the A-weighting correction
' Args:     None
' Comments: (1) Could we do this a simpler way? <--TODO: upgrade for Trace 3.0
'==============================================================================

Sub A_weight_Corrections()


'If T_SheetType = "OCT" Then
'    For col = 2 To 13
'        Select Case col
'        Case 2
'        Cells(T_FreqRow, col).Value = "A Weighting"
'        Case 5
'        Cells(T_FreqRow, col).Value = -39.4
'        Case 6
'        Cells(T_FreqRow, col).Value = -26.2
'        Case 7
'        Cells(T_FreqRow, col).Value = -16.1
'        Case 8
'        Cells(T_FreqRow, col).Value = -8.6
'        Case 9
'        Cells(T_FreqRow, col).Value = -3.2
'        Case 10
'        Cells(T_FreqRow, col).Value = 0#
'        Case 11
'        Cells(T_FreqRow, col).Value = -1.2
'        Case 12
'        Cells(T_FreqRow, col).Value = 1#
'        Case 13
'        Cells(T_FreqRow, col).Value = -1.1
'        End Select
'    Next col
'ElseIf T_SheetType = "OCTA" Then
'    For col = 2 To 13
'        Select Case col
'        Case 2
'        Cells(T_FreqRow, col).Value = "A Weighting"
'        Case 5
'        Cells(T_FreqRow, col).Value = 39.4
'        Case 6
'        Cells(T_FreqRow, col).Value = 26.2
'        Case 7
'        Cells(T_FreqRow, col).Value = 16.1
'        Case 8
'        Cells(T_FreqRow, col).Value = 8.6
'        Case 9
'        Cells(T_FreqRow, col).Value = 3.2
'        Case 10
'        Cells(T_FreqRow, col).Value = 0#
'        Case 11
'        Cells(T_FreqRow, col).Value = 1.2
'        Case 12
'        Cells(T_FreqRow, col).Value = 1#
'        Case 13
'        Cells(T_FreqRow, col).Value = 1.1
'        End Select
'    Next col
'ElseIf T_SheetType = "TO" Then
' For col = 2 To 25
'        Select Case col
'        Case 2
'        Cells(T_FreqRow, col).Value = "A Weighting"
'        Case 5
'        Cells(T_FreqRow, col).Value = -30.2
'        Case 6
'        Cells(T_FreqRow, col).Value = -26.2
'        Case 7
'        Cells(T_FreqRow, col).Value = -22.5
'        Case 8
'        Cells(T_FreqRow, col).Value = -19.1
'        Case 9
'        Cells(T_FreqRow, col).Value = -16.1
'        Case 10
'        Cells(T_FreqRow, col).Value = -13.4
'        Case 11
'        Cells(T_FreqRow, col).Value = -10.9
'        Case 12
'        Cells(T_FreqRow, col).Value = -8.6
'        Case 13
'        Cells(T_FreqRow, col).Value = -6.6
'        Case 14
'        Cells(T_FreqRow, col).Value = -4.8
'        Case 15
'        Cells(T_FreqRow, col).Value = -3.2
'        Case 16
'        Cells(T_FreqRow, col).Value = -1.9
'        Case 17
'        Cells(T_FreqRow, col).Value = -0.8
'        Case 18
'        Cells(T_FreqRow, col).Value = 0#
'        Case 19
'        Cells(T_FreqRow, col).Value = 0.6
'        Case 20
'        Cells(T_FreqRow, col).Value = 1#
'        Case 21
'        Cells(T_FreqRow, col).Value = 1.2
'        Case 22
'        Cells(T_FreqRow, col).Value = 1.3
'        Case 23
'        Cells(T_FreqRow, col).Value = 1.2
'        Case 24
'        Cells(T_FreqRow, col).Value = 1#
'        Case 25
'        Cells(T_FreqRow, col).Value = 0.5
'        End Select
'    Next col
'    ElseIf T_SheetType = "TOA" Then
'    For col = 2 To 25
'        Select Case col
'        Case 2
'        Cells(T_FreqRow, col).Value = "A Weighting"
'        Case 5
'        Cells(T_FreqRow, col).Value = 30.2
'        Case 6
'        Cells(T_FreqRow, col).Value = 26.2
'        Case 7
'        Cells(T_FreqRow, col).Value = 22.5
'        Case 8
'        Cells(T_FreqRow, col).Value = 19.1
'        Case 9
'        Cells(T_FreqRow, col).Value = 16.1
'        Case 10
'        Cells(T_FreqRow, col).Value = 13.4
'        Case 11
'        Cells(T_FreqRow, col).Value = 10.9
'        Case 12
'        Cells(T_FreqRow, col).Value = 8.6
'        Case 13
'        Cells(T_FreqRow, col).Value = 6.6
'        Case 14
'        Cells(T_FreqRow, col).Value = 4.8
'        Case 15
'        Cells(T_FreqRow, col).Value = 3.2
'        Case 16
'        Cells(T_FreqRow, col).Value = 1.9
'        Case 17
'        Cells(T_FreqRow, col).Value = 0.8
'        Case 18
'        Cells(T_FreqRow, col).Value = 0#
'        Case 19
'        Cells(T_FreqRow, col).Value = -0.6
'        Case 20
'        Cells(T_FreqRow, col).Value = -1#
'        Case 21
'        Cells(T_FreqRow, col).Value = -1.2
'        Case 22
'        Cells(T_FreqRow, col).Value = -1.3
'        Case 23
'        Cells(T_FreqRow, col).Value = -1.2
'        Case 24
'        Cells(T_FreqRow, col).Value = -1#
'        Case 25
'        Cells(T_FreqRow, col).Value = -0.5
'        End Select
'    Next col
'Else
'ErrorTypeCode
'End If

End Sub

'==============================================================================
' Name:     MoveUp
' Author:   PS
' Desc:     Moves selected calculation row(s) up by 1.
' Args:     None
' Comments: (1) Should maintain links, but may break data validation?
'==============================================================================

Sub MoveUp()
Dim StartRw As Integer
Dim EndRw As Integer

Application.ScreenUpdating = False

StartRw = Selection.Row
EndRw = Selection.Row + Selection.Rows.Count - 1

    'check for adjacent empty row
    If Cells(StartRw - 1, 2).Value <> "" Then
    msg = MsgBox("There appears to be data in the cells above.  " & _
        "Continuing will delete this data. Do you want to continue?", _
        vbYesNo, "Check yo'self")
        If msg = vbNo Then End
    End If

Range("B" & StartRw & ":D" & EndRw).Cut Destination:=Range("B" & StartRw - 1 & ":D" & EndRw - 1) 'Description
    'TODO: update for Trace 3
    If Left(T_SheetType, 3) = "OCT" Then
    Range("E" & StartRw & ":O" & EndRw).Cut Destination:=Range("E" & StartRw - 1 & ":O" & EndRw - 1) 'Formulas
    Range("B" & StartRw - 1 & ":O" & StartRw - 1).Copy 'formats
    Range("B" & StartRw & ":O" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf Left(T_SheetType, 2) = "TO" Then
    Range("E" & StartRw & ":AA" & EndRw).Cut Destination:=Range("E" & StartRw - 1 & ":AA" & EndRw - 1) 'Formulas
    Range("B" & StartRw - 1 & ":AA" & StartRw - 1).Copy 'formats
    Range("B" & StartRw & ":AA" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf T_SheetType = "LF_TO" Then
    Range("E" & StartRw & ":AG" & EndRw).Cut Destination:=Range("E" & StartRw - 1 & ":AG" & EndRw - 1) 'Formulas
    Range("B" & StartRw - 1 & ":AG" & StartRw - 1).Copy 'formats
    Range("B" & EndRw & ":AG" & EndRw).PasteSpecial Paste:=xlPasteFormats
    End If
    
'move to select lower row
Range(Cells(StartRw - 1, 2), Cells(EndRw - 1, 2)).Select

Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub

'==============================================================================
' Name:     MoveDown
' Author:   PS
' Desc:     Moves selected calculation row(s) down by 1.
' Args:     None
' Comments: (1) Should maintain links, but may break data validation?
'==============================================================================

Sub MoveDown()
Dim StartRw As Integer
Dim EndRw As Integer

Application.ScreenUpdating = False

StartRw = Selection.Row
EndRw = Selection.Row + Selection.Rows.Count - 1

    'check for adjacent empty row
    If Cells(EndRw + 1, 2).Value <> "" Then
    msg = MsgBox("There appears to be data in the cells above.  " & _
        "Continuing will delete this data. Do you want to continue?", _
        vbYesNo, "Check yo'self")
        If msg = vbNo Then End
    End If

Range("B" & StartRw & ":D" & EndRw).Cut Destination:=Range("B" & StartRw + 1 & ":D" & EndRw + 1) 'Description
    'TODO: update for Trace 3
    If Left(T_SheetType, 3) = "OCT" Then
    Range("E" & StartRw & ":O" & EndRw).Cut Destination:=Range("E" & StartRw + 1 & ":O" & EndRw + 1) 'Formulas
    Range("B" & StartRw + 1 & ":O" & StartRw + 1).Copy 'formats
    Range("B" & StartRw & ":O" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf Left(T_SheetType, 2) = "TO" Then
    Range("E" & StartRw & ":AA" & EndRw).Cut Destination:=Range("E" & StartRw + 1 & ":AA" & EndRw + 1) 'Formulas
    Range("B" & StartRw + 1 & ":AA" & StartRw + 1).Copy 'formats
    Range("B" & StartRw & ":AA" & StartRw).PasteSpecial Paste:=xlPasteFormats
    ElseIf T_SheetType = "LF_TO" Then
    Range("E" & StartRw & ":AG" & EndRw).Cut Destination:=Range("E" & StartRw + 1 & ":AG" & EndRw + 1) 'Formulas
    Range("B" & StartRw + 1 & ":AG" & StartRw + 1).Copy 'formats
    Range("B" & StartRw & ":AG" & StartRw).PasteSpecial Paste:=xlPasteFormats
    End If

'move to select lower row
Range(Cells(StartRw + 1, 2), Cells(EndRw + 1, 2)).Select

Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub

 '==============================================================================
' Name:     MoveLeft
' Author:   PS
' Desc:     Moves selected spectrum one band to the left
' Args:     None
' Comments: (1) Applies to main columns only, descriptions stay put
'==============================================================================

Sub MoveLeft()
Dim StartRw As Integer
Dim EndRw As Integer
Dim startCol As Integer
Dim endCol As Integer

Application.ScreenUpdating = False

StartRw = Selection.Row
EndRw = Selection.Row + Selection.Rows.Count - 1
'startCol = GetSheetTypeColumns(SheetType, "LossGainStart")
'endCol = GetSheetTypeColumns(SheetType, "LossGainEnd")

Range(Cells(StartRw, T_LossGainStart + 1), Cells(EndRw, T_LossGainEnd)).Copy
Cells(StartRw, T_LossGainStart).PasteSpecial Paste:=xlPasteValues
Cells(StartRw, T_LossGainEnd).ClearContents

Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub

 '==============================================================================
' Name:     MoveRight
' Author:   PS
' Desc:     Moves selected spectrum one band to the right
' Args:     None
' Comments: (1) Applies to main columns only, descriptions stay put
'==============================================================================

Sub MoveRight()
Dim StartRw As Integer
Dim EndRw As Integer
Dim startCol As Integer
Dim endCol As Integer

Application.ScreenUpdating = False

StartRw = Selection.Row
EndRw = Selection.Row + Selection.Rows.Count - 1
'startCol = GetSheetTypeColumns(SheetType, "LossGainStart")
'endCol = GetSheetTypeColumns(SheetType, "LossGainEnd")

Range(Cells(StartRw, T_LossGainStart), Cells(EndRw, T_LossGainEnd - 1)).Copy
Cells(StartRw, T_LossGainStart + 1).PasteSpecial Paste:=xlPasteValues
Cells(EndRw, T_LossGainStart).ClearContents

Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub

 '==============================================================================
' Name:     Row Reference
' Author:   PS
' Desc:     Puts in a reference to data in another row or sheet
' Args:     None
' Comments: (1) Updated to work across sheets
'==============================================================================

Sub RowReference()
Dim FirstRow As Integer
Dim LastRow As Integer
Dim SheetName As String

frmRowReference.Show

    If btnOkPressed = False Then
    End
    End If

    If UserSelectedAddress = "" Then End 'error catch
    
    SheetName = GetSheetName(UserSelectedAddress)
    FirstRow = GetFirstRow(UserSelectedAddress)
    LastRow = GetLastRow(UserSelectedAddress)
    
    'TODO: check for mismatched sheet types
    'something like:
'    If Sheets(SheetName).Range("TYPECODE") <> T_SheetType Then
'
'    End If

    If LookupMultiRow = False Then
    Cells(Selection.Row, T_Description).Value = "=CONCAT(""Ref: ""," & _
        SheetName & "$B$" & FirstRow & ")"
    Cells(Selection.Row, T_LossGainStart).Value = "=" & _
        SheetName & "E$" & FirstRow
    ExtendFunction
    
    Else 'multimode = true
    SetDataValidation T_Description, "=" & SheetName & "$B$" & FirstRow & _
        ":$B$" & LastRow
        
    'select first entry by default
    Cells(Selection.Row, T_Description).Value = Range(SheetName & "$B$" & FirstRow)
    
        'create index-match formula <--TODO: update this for Trace 3
        If T_SheetType = "OCT" Then
        'Debug.Print "=INDEX(" & SheetName & "$E$" & FirstRow & ":$M$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
        "," & SheetName & "$B$" & FirstRow & ":$B$" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!" & T_FreqStartRng & "," & SheetName & "$" & T_FreqStartRng & ":$M$6,0))"
        Cells(Selection.Row, T_LossGainStart).Value = "=INDEX(" & SheetName & "$E$" & FirstRow & ":$M$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
        "," & SheetName & "$B$" & FirstRow & ":$B" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!" & T_FreqStartRng & "," & SheetName & "$" & T_FreqStartRng & ":$M$6,0))" '<----note that SheetName includes apostrophe character and ActiveSheet.Name does not.....trickyyyyy
        ExtendFunction
        
        ElseIf Left(T_SheetType, 2) = "TO" Then
        Cells(Selection.Row, T_LossGainStart).Value = "=INDEX(" & SheetName & "$E$" & FirstRow & ":$Y$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
        "," & SheetName & "$B$" & FirstRow & ":$B" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!" & T_FreqStartRng & "," & SheetName & "$" & T_FreqStartRng & ":$Y$6,0))"
        ExtendFunction
        
        ElseIf T_SheetType = "MECH" Then
        Cells(Selection.Row, T_RegenStart).Value = "=INDEX(" & SheetName & "$T$" & FirstRow & ":$AB$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
        "," & SheetName & "$B$" & FirstRow & ":$B" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!T$6," & SheetName & "$T$6:$AB$6,0))"
        ExtendFunction (True)
        End If
    
    End If
    
'apply style
SetTraceStyle "Reference"

End Sub

'==============================================================================
' Name:     ToggleActive
' Author:   AA
' Desc:     Toggles cell contents in a single row within the working area
'           between active and inactive. Preserves original formula and format.
' Args:     None
' Comments: (1)
'==============================================================================

Sub ToggleActive()
Dim startCol As Integer, endCol As Integer, WrkRow As Integer
Dim CharKeep As Integer
Dim Orig As String, OrigFmt As String, FormatArchive As String
Dim NewValue As String

WrkRow = Selection.Row

'CHECK FOR NON HEADER ROWS
CheckTemplateRow (WrkRow)

'    ' Set start and end columns based on sheet type
'    If TypeCode = "OCT" Or TypeCode = "OCTA" Then
'        ' Redundant code for later use [TypeCode = "CT" Or TypeCode = "P2W" Or
'        ' TypeCode = "R2R" Or TypeCode = "RN"]
'        startCol = 5
'        endCol = 13
'    ElseIf TypeCode = "LF_TO" Then
'        startCol = 5
'        endCol = 31
'    ElseIf TypeCode = "TO" Or TypeCode = "TOA" Then
'        'Redundant code for later use [Or TypeCode = "GLZ" Or TypeCode = "SII"]
'        startCol = 5
'        endCol = 25
'    'ElseIf TypeCode = "RT" Then
'    '    StartCol = 3
'    '    EndCol = 3
'    Else
'        ErrorTypeCode
'    End If


    ' Toggles between active and inactive
    For i = T_LossGainStart To T_LossGainEnd
    
        ' If currently inactive, toggle active. Looks for [,"\*] within formula.
        If InStr(1, Cells(WrkRow, i).Formula, ",""\*") <> 0 Then
            Orig = Cells(WrkRow, i).Formula
            Dim Pos1 As Integer, Pos2_1 As Integer
            Pos1 = 7
            Pos2_1 = InStr(1, Orig, ",""\*") - Pos1 'looks for position of [,"\*]
            NewValue = Mid(Orig, Pos1, Pos2_1)
            
            ' Whether original cell value was a formula or not
            If Mid(Orig, InStr(1, Orig, "T(N(") + 4, 1) = 1 Then
                Cells(WrkRow, i).Value = NewValue
            Else
                Cells(WrkRow, i).Value = "=" & NewValue
            End If
                
        ' If cell is not empty, toggle inactive
        ElseIf IsEmpty(Cells(WrkRow, i)) = False Then
            
            ' If cell is not empty, and is a formula
            If Left(Cells(WrkRow, i).Formula, 1) = "=" Then
                Orig = Cells(WrkRow, i).Formula
                CharKeep = 2
            
            ' If cell is not empty, but isn't a formula
            Else
                Orig = Cells(WrkRow, i).Value
                CharKeep = 1
                
            End If
            ' Take current cell value/formula and rework to new commented formula
            OrigFmt = Cells(WrkRow, i).NumberFormat
            FormatArchive = Replace(OrigFmt, """", """""")
            FormatArchive = "\*" & FormatArchive & "\*;\*-" & FormatArchive _
                & "\*;\*" & FormatArchive & "\*"
            
            ' The following is a janky work-around way of keeping information
            ' within the cell formula. Arg [CharKeep] has the unintended side
            ' effect of storing whether the original cell was a formula or a
            ' simple value. The "&T(N(xxx))" method used store the value in the
            ' formula, and is such a roundabout method because the displayed cell
            ' result is text, not a number.
            Cells(WrkRow, i).Value = "=TEXT(" & Mid(Orig, CharKeep) & ",""" & _
                FormatArchive & """)" & "&T(N(" & CharKeep & "))"
                
        End If
        
    Next i

End Sub

'==============================================================================
' Name:     SingleCorrection
' Author:   PS
' Desc:     Inserts a single correction into the first parameter column,
'           referred to all octave bands
' Args:     None
' Comments: (1)
'==============================================================================
Sub SingleCorrection()
'Dim col As Integer

Cells(Selection.Row, T_Description).Value = "Correction"
Cells(Selection.Row, T_LossGainStart).Value = "=" & T_ParamRng(0)
Cells(Selection.Row, T_ParamStart).Value = -5
ParameterMerge (Selection.Row)
SetUnits "dB", T_ParamStart, 0
ExtendFunction
SetTraceStyle "Input", True
Cells(Selection.Row, T_ParamStart).Select 'move to parameter column t set value
End Sub


'==============================================================================
' Name:     AutoSum
' Author:   PS
' Desc:     Sums all rows until a blank row is reached.
' Args:     ApplyStyleCode - String that says what style to use
'           LineDescStr - String for the T_description column
' Comments: (1)
'==============================================================================
Sub AutoSum(Optional ApplyStyleCode As String, Optional LineDescStr As String)
Dim FindRw As Integer
Dim ScanCol As Integer
Dim FoundRw As Boolean

    If LineDescStr = "" Then
    Cells(Selection.Row, T_Description).Value = "Total"
    Else
    Cells(Selection.Row, T_Description).Value = LineDescStr
    End If

'find end of range
FindRw = Selection.Row - 1 'one above findrw
ScanCol = Selection.Column
foudnRw = False

    While FoundRw = False
    FindRw = FindRw - 1
    
        If FindRw < 8 Then 'A weighting is on line 7 for all template sheets
        FindRw = 7 'A weighting line is the same as a blank line
        FoundRw = True
        End If
        
        If Cells(FindRw, ScanCol).Value = "" Then FoundRw = True
        
    Wend
    
Cells(Selection.Row, T_LossGainStart).Value = "=SUM(" & _
    Range(Cells(FindRw + 1, T_LossGainStart), _
    Cells(Selection.Row - 1, T_LossGainStart)).Address(False, False) & ")"
    
ExtendFunction
    
    'Limit the options to the three main styles
    If ApplyStyleCode = "Total" Then
    SetTraceStyle "Total"
    ElseIf ApplyStyleCode = "Subtotal" Then
    SetTraceStyle "Subtotal"
    ElseIf ApplyStyleCode = "Normal" Then
    SetTraceStyle "Normal"
    Else 'default to Subtotal
    SetTraceStyle "Subtotal"
    End If
    
End Sub

'==============================================================================
' Name:     Manual_ExtendFunction
' Author:   PS
' Desc:     Extends function from column E to the full range - user initiated.
' Args:     None
' Comments: (1) ExtendFunction is often called by other functions, this is just
'           an allowance for users to do the same.
'==============================================================================
Sub Manual_ExtendFunction()
SetSheetTypeControls
ExtendFunction
End Sub

'==============================================================================
' Name:     OneThirdsToOctave
' Author:   PS
' Desc:     Converts one-third octave bands into octave bands. Functions are
'           inserted for logarithmic sum or av, if required.
' Args:     None
' Comments: (1)
'==============================================================================
Sub OneThirdsToOctave()

Dim splitAddr() As String
Dim SheetName As String
Dim rw As Integer
Dim refCol As Integer
Dim targetRange As String


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Left(T_SheetType, 3) = "OCT" Then 'OCT or OCTA sheet types
    
    frmConvert.Show

        If btnOkPressed = False Then
        End
        End If

    If UserSelectedAddress = "" Then End 'error catch
    
    SheetName = GetSheetName(UserSelectedAddress)
    FirstRow = GetFirstRow(UserSelectedAddress)
    
    Cells(Selection.Row, T_Description).Value = "Convert from TO"
    Cells(Selection.Row, T_Comment).Value = "=" & SheetName & "$B" & FirstRow
    
    If UserSelectedAddress = "" Then End
    
    splitAddr = Split(UserSelectedAddress, "$", Len(UserSelectedAddress), vbTextCompare)
    
    SheetName = splitAddr(LBound(splitAddr)) 'sheet is the first element
    rw = CInt(splitAddr(UBound(splitAddr))) 'row is the last element
    
        refCol = 5
        For col = T_LossGainStart To T_LossGainEnd
        targetRange = Range(Cells(rw, refCol), Cells(rw, refCol + 2)).Address(False, False)
            Select Case SumAverageMode 'selected from radio boxes in form frmConvert
            Case Is = "Sum"
            Cells(Selection.Row, col).Value = "=SPLSUM(" & SheetName & targetRange & ")"
            Case Is = "Average"
            Cells(Selection.Row, col).Value = "=AVERAGE(" & SheetName & targetRange & ")"
            Case Is = "Log Av"
            Cells(Selection.Row, col).Value = "=SPLAV(" & SheetName & targetRange & ")"
            Case Is = "TL" 'positive spectra returned as positive
            Cells(Selection.Row, col).Value = "=TL_ThirdsToOctave(" & SheetName & targetRange & ")"
            End Select
        refCol = refCol + 3
        Next col
        
    'apply reference style
    SetTraceStyle "Reference"
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf T_SheetType = "CVT" Then
    
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
        For col = T_RegenStart To T_RegenEnd
        targetRange = Range(Cells(rw, refCol), Cells(rw, refCol + 2)).Address(False, False)
            Select Case SumAverageMode 'selected from radio boxes in form frmConvert
            Case Is = "Sum"
            Cells(Selection.Row, col).Value = "=SPLSUM(" & SheetName & targetRange & ")"
            Case Is = "Average"
            Cells(Selection.Row, col).Value = "=AVERAGE(" & SheetName & targetRange & ")"
            Case Is = "Log Av"
            Cells(Selection.Row, col).Value = "=SPLAV(" & SheetName & targetRange & ")"
            Case Is = "TL" 'positive spectra returned as positive
            Cells(Selection.Row, col).Value = "=TL_ThirdsToOctave(" & SheetName & targetRange & ")"
            End Select
        refCol = refCol + 3
        Next col
        
    'apply reference style ''''''''''''''maybe dont for CVT sheet?
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf Left(T_SheetType, 2) = "TO" Then 'TO or TOA
    msg = MsgBox("Imports into OCT or OCTA", vbOKOnly, "WRONG WAY GO BACK!")
    Else
    ErrorTypeCode
    End If

End Sub

'==============================================================================
' Name:     ConvertAWeight
' Author:   PS
' Desc:     Adds the A-weighting from the top of the sheet to to the current
'           row making it an C-weighted spectrum shape.
' Args:     None
' Comments: (1) Sheet currently supports OCT, OCTA, TO, TOA sheets
'==============================================================================
Sub ConvertAWeight()

    'screen for sheet types and set description
    If Left(T_SheetType, 3) = "OCT" Or Left(T_SheetType, 2) = "TO" Then
        If Right(T_SheetType, 1) = "A" Then 'a-weighted sheets
        Cells(Selection.Row, T_Description).Value = _
            Cells(Selection.Row - 1, T_Description).Value & " (Linear)"
        Else
        Cells(Selection.Row, T_Description).Value = _
            Cells(Selection.Row - 1, T_Description).Value & " (A Weighted)"
        End If
    Else
    ErrorOCTTOOnly
    End If
    
'build formula
Cells(Selection.Row, T_LossGainStart).Value = "=" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & "+" & _
    Cells(7, T_LossGainStart).Address(True, False)
ExtendFunction

End Sub

'==============================================================================
' Name:     ConvertCWeight
' Author:   PS
' Desc:     Adds the C-weighting curve in a row, then adds it to the current
'           row making it an C-weighted spectrum shape.
' Args:     None
' Comments: (1) Sheet currently supports OCT, OCTA, TO, TOA sheets
'==============================================================================
Sub ConvertCWeight()

ErrorDoesNotExist 'TODO: implement this!
'hint, build a c-weight curve function first

'    'screen for sheet types and set description
'    If Left(T_SheetType, 3) = "OCT" Or Left(T_SheetType, 2) = "TO" Then
'        If Right(T_SheetType, 1) = "A" Then 'a-weighted sheets
'        Cells(Selection.Row, T_Description).Value = _
'            Cells(Selection.Row - 1, T_Description).Value & " (Linear)"
'        Else
'        Cells(Selection.Row, T_Description).Value = _
'            Cells(Selection.Row - 1, T_Description).Value & " (A Weighted)"
'        End If
'    Else
'    ErrorOCTTOOnly
'    End If
'
'
''build formula
'Cells(Selection.Row, T_LossGainStart).Value = "=" & _
'    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & "+" & _
'    Cells(7, T_LossGainStart).Address(True, False)
'ExtendFunction

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Called from other functions, no check needed from this point onwards
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     CVT_OneThirds2Oct
' Author:   PS
' Desc:     Ports to new function OneThirdsToOctave
' Args:     None
' Comments: (1)
'==============================================================================
Sub CVT_OneThirds2Oct()
    If IsNamedRange("TYPECODE") Then 'send to normal conversion form
    SetSheetTypeControls
    OneThirdsToOctave
    End If
End Sub

'==============================================================================
' Name:     ExtendFunction
' Author:   PS
' Desc:     Copies formulas to the correct ranges
' Args:     None
' Comments: (1) Sheet currently supports OCT, OCTA, TO, TOA sheets
'==============================================================================
Sub ExtendFunction(Optional ApplyToRegen As Boolean)
Dim StartAddr As String
StartAddr = Selection.Address
    If T_LossGainStart < 1 Or T_LossGainEnd < 1 Then
    'public variables not defined, let's fix that
    SetSheetTypeControls
    End If
    
    If ApplyToRegen = True Then
    Cells(Selection.Row, T_RegenStart).Copy
    Range(Cells(Selection.Row, T_RegenStart), Cells(Selection.Row, T_RegenEnd)).PasteSpecial (xlPasteFormulas)
    Else
    Cells(Selection.Row, T_LossGainStart).Copy
    Range(Cells(Selection.Row, T_LossGainStart), Cells(Selection.Row, T_LossGainEnd)).PasteSpecial (xlPasteFormulas)
    End If
    
Application.CutCopyMode = False
Range(StartAddr).Select
End Sub

'==============================================================================
' Name:     ParameterMerge
' Author:   PS
' Desc:     Merges parameter columns
' Args:     rw - row number
'           NumCols - defaults to 2, can change when requested
' Comments: (1) Neat.
'==============================================================================
Sub ParameterMerge(rw As Integer, Optional NumCols As Integer)
'<-TODO update this function for preset sheet types
    If IsMissing(NumCols) Or NumCols < 2 Then NumCols = 2
    
Range(T_ParamRng(0), T_ParamRng(NumCols - 1)).Merge
Range(T_ParamRng(0)).HorizontalAlignment = xlCenter
Range(T_ParamRng(0)).VerticalAlignment = xlCenter

End Sub

'==============================================================================
' Name:     ParameterMerge
' Author:   PS
' Desc:     Unmerges parameter columns, sets borders
' Args:     rw - row number
' Comments: (1) Neat.
'==============================================================================
Sub ParameterUnmerge(rw As Integer)

Range(Cells(rw, T_ParamStart), Cells(rw, T_ParamEnd)).UnMerge

    With Range(Cells(rw, T_ParamStart), Cells(rw, T_ParamEnd))
    .Borders.LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Borders(xlInsideVertical).Weight = xlHairline
    End With

End Sub

'==============================================================================
' Name:     Sparkline
' Author:   PS
' Desc:     Puts sparkline on current row
' Args:     None
' Comments: (1)
'==============================================================================
Sub Sparkline()
ParameterMerge (Selection.Row)
CreateSparkline Selection.Row, 0
End Sub

'==============================================================================
' Name:     CreateSparkline
' Author:   PS
' Desc:     Makes Sparkline for the spectrum
' Args:     rw - row number
'           p_index - parameter index number, usually 0 to 3
' Comments: (1)
'==============================================================================
Sub CreateSparkline(rw As Integer, p_index As Integer, Optional colorIndex As Integer)

Dim DataRangeStr As String

    If p_index <= UBound(T_ParamRng) Then
        'Set where the data to be graphed is
        If T_SheetType = "MECH" Then
        DataRangeStr = Range(Cells(rw, T_RegenStart), _
            Cells(rw, T_RegenEnd)).Address
        Else
        DataRangeStr = Range(Cells(rw, T_LossGainStart), _
            Cells(rw, T_LossGainEnd)).Address
        End If
        
    'add the sparkline
    Cells(rw, T_ParamStart + p_index).SparklineGroups.Add _
        Type:=xlSparkLine, SourceData:=DataRangeStr
        
        'formatting etc
        With Cells(rw, T_ParamStart + p_index).SparklineGroups.Item(1)
        .SeriesColor.Color = colorIndex 'defaults to 0
        .SeriesColor.TintAndShade = 0
        End With
        
    End If
End Sub

'==============================================================================
' Name:     SelectNextRow
' Author:   PS
' Desc:     Selects the next row
' Args:     None
' Comments: (1) It's just neater this way
'==============================================================================
Sub SelectNextRow()
Cells(Selection.Row + 1, Selection.Column).Select 'move down
SetSheetTypeControls 'update variables
End Sub




'==============================================================================
' Name:     Summarise_RT
' Author:   PS
' Desc:     Summarises RT times
' Args:     None
' Comments: (1) still fairly hacky. TODO: make less hacky
'==============================================================================
Sub Summarise_RT()
Dim WriteRw As Integer
Dim FormulaStr As String
Dim FinishesStr As String

WriteRw = Selection.Row 'start from here

    For sh = 2 To Sheets.Count
    Sheets(sh).Activate
        'check if typecode is RT
        If IsNamedRange("RT") Then
        
        'hyperlink
        FormulaStr = "=HYPERLINK(""#'" & Sheets(sh).Name & "'!A1"",""" & _
            Sheets(sh).Name & """)"
            'Debug.Print FormulaStr
        
        Sheets("SUMMARY").Cells(WriteRw, T_Description).Value = FormulaStr
            
            'Values
            For col = 2 To 7
            Sheets("SUMMARY").Cells(WriteRw, col + 5).Value = _
                "='" & Sheets(sh).Name & "'!" & Cells(29, col + 5).Address
            Sheets("SUMMARY").Cells(WriteRw, col + 5).NumberFormat = "0.0"
            Next col
            
            'finishes
            FinishesStr = ""
            For materialrw = 13 To 24
                If Cells(materialrw, 3).Value > 0 Then
                FinishesStr = FinishesStr & ";" & Cells(materialrw, 5).Value
                End If
            Next materialrw
            'Debug.Print FinishesStr
            Sheets("SUMMARY").Cells(WriteRw, 16).Value = FinishesStr
            
        WriteRw = WriteRw + 1
        End If
    Next sh
Sheets("SUMMARY").Activate
End Sub


'**************
'Code Graveyard
'**************


'Sub ClearSparkline(ParamCol As Integer)
'    If ParamCol <= UBound(T_ParamRng) Then
'    Cells(Selection.Row, T_ParamStart + ParamCol).SparklineGroups.Clear
'    'Range(T_ParamRng(ParamCol)).SparklineGroups.Clear
'    End If
'End Sub

'Sub UserInputFormat(SheetType As String)
'    If Left(SheetType, 3) = "OCT" Then
'    Range(Cells(Selection.Row, 5), Cells(Selection.Row, 13)).Interior.Color = RGB(251, 251, 143)
'    ElseIf Left(SheetType, 2) = "TO" Then
'    Range(Cells(Selection.Row, 5), Cells(Selection.Row, 25)).Interior.Color = RGB(251, 251, 143)
'    Else 'do nothing
'    End If
'End Sub

'Sub UserInputFormat_ParamCol(SheetType As String) 'legacy code, will redirect to new function
'
'fmtUserInput (SheetType)
'
''OLD VERSION, USES COLOURS, NOT STYLES
''    If Left(SheetType, 3) = "OCT" Then
''    Range(Cells(Selection.Row, 14), Cells(Selection.Row, 15)).Interior.Color = RGB(251, 251, 143)
''    ElseIf Left(SheetType, 2) = "TO" Then
''    Range(Cells(Selection.Row, 26), Cells(Selection.Row, 27)).Interior.Color = RGB(251, 251, 143)
''    ElseIf SheetType = "LF_TO" Then
''    Range(Cells(Selection.Row, 32), Cells(Selection.Row, 33)).Interior.Color = RGB(251, 251, 143)
''    Else
''    SheetTypeUnknownError(SheetType)
''    End If
'
'End Sub

'Sub ErrorTypeCode() '<TODO:  is this needed?
'msg = MsgBox("Not implemented for Typecode: " & T_SheetType, vbOKOnly, "Error - Sheet Type")
'End
'End Sub

