Attribute VB_Name = "RowOperations"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
Public UserSelectedAddress As String
Public SumAverageMode As String
Public LookupMultiRow As Boolean
Public RegenDestinationRange As Boolean

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
Dim Rw As Integer
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
    For Rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
    
    'Description
    Cells(Rw, T_Description).ClearContents
    Cells(Rw, T_Description).ClearComments
    Cells(Rw, T_Description).Validation.Delete 'for dropdown boxes
        
    'LossGain values
    Range(Cells(Rw, T_LossGainStart), Cells(Rw, T_LossGainEnd)).ClearContents
        
        'Regenerated noise columns
        If T_RegenStart <> -1 Then
        Range(Cells(Rw, T_RegenStart), Cells(Rw, T_RegenEnd)).ClearContents
        End If
       
        'Comment
        If T_Comment <> -1 Then
        Cells(Rw, T_Comment).ClearContents
        End If
    
        'Parameter columns
        If T_ParamStart >= 0 Then
        ParameterUnmerge (Rw)
            For P = T_ParamStart To T_ParamEnd
            Cells(Rw, P).Validation.Delete
            Cells(Rw, P).ClearContents
            Cells(Rw, P).ClearComments
            Cells(Rw, P).NumberFormat = "General"
            Next P
            
            'sparklines
            For i = 0 To 3
            Cells(Rw, T_ParamStart + i).SparklineGroups.Clear
            Next i
        End If
        
    'remove heatmap
    Range(Cells(Rw, T_Description), Cells(Rw, T_LossGainEnd)).FormatConditions.Delete
    
    'apply styles and reset units
    ApplyTraceStyle "Trace Normal", Rw
    SetUnits "Clear", T_LossGainStart, 0, T_LossGainEnd
    'styles for parameter columns
        If T_ParamStart >= 0 Then
        ApplyTraceStyle "Trace Normal", Rw, True
        End If
        
    'bold to overall column
    Cells(Rw, T_LossGainStart - 1).Font.Bold = True
    
    'lock cells
    Range(Cells(Rw, T_Description), Cells(Rw, T_LossGainEnd)).Locked = True
    
    Next Rw
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
Dim Rw As Integer
Dim startCol As Integer
Dim endCol As Integer

    For Rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
        For Col = T_LossGainStart To T_LossGainEnd
            If Cells(Rw, Col).HasFormula Then
                'check if second character of formula is minus
                If Mid(Cells(Rw, Col).Formula, 2, 1) = "-" Then
                Cells(Rw, Col).Formula = Replace(Cells(Rw, Col).Formula, _
                    "=-", "=", 1, Len(Cells(Rw, Col).Formula), vbTextCompare)
                Else
                Cells(Rw, Col).Formula = Replace(Cells(Rw, Col).Formula, _
                    "=", "=-", 1, Len(Cells(Rw, Col).Formula), vbTextCompare)
                End If
            Else
                Cells(Rw, Col).Value = Cells(Rw, Col).Value * -1
            End If
        Next Col
    Next Rw

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
Dim DestinationCol As Integer

    If T_SheetType = "MECH" Then
    frmRowReference.optRegenSWL.Enabled = True
    Else
    frmRowReference.optRegenSWL.Enabled = False
    End If

frmRowReference.Show

    'error catch
    If btnOkPressed = False Then End
    If UserSelectedAddress = "" Then End
    
SheetName = GetSheetName(UserSelectedAddress)
FirstRow = GetFirstRow(UserSelectedAddress)
LastRow = GetLastRow(UserSelectedAddress)

    'set destination column
    If RegenDestinationRange = True Then
    DestinationCol = T_RegenStart
    Else
    DestinationCol = T_LossGainStart
    End If

    'TODO: check for mismatched sheet types
    'something like:
'    If Sheets(SheetName).Range("TYPECODE") <> T_SheetType Then
'
'    End If

    If LookupMultiRow = False Then
    SetDescription "=CONCAT(""Ref: ""," & SheetName & "$B$" & FirstRow & ")"
    Cells(Selection.Row, DestinationCol).Value = "=" & _
        SheetName & "E$" & FirstRow
    ExtendFunction (RegenDestinationRange)
    
    Else 'multimode = true
    SetDataValidation T_Description, "=" & SheetName & "$B$" & FirstRow & _
        ":$B$" & LastRow
        
    'select first entry by default
    SetDescription Range(SheetName & "$B$" & FirstRow)
    
        'create index-match formula <--TODO: update this for Trace 3
        If T_SheetType = "OCT" Then
        'Debug.Print "=INDEX(" & SheetName & "$E$" & FirstRow & ":$M$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
        "," & SheetName & "$B$" & FirstRow & ":$B$" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!" & T_FreqStartRng & "," & SheetName & "$" & T_FreqStartRng & ":$M$6,0))"
        Cells(Selection.Row, DestinationCol).Value = "=INDEX(" & SheetName & "$E$" & FirstRow & ":$M$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
        "," & SheetName & "$B$" & FirstRow & ":$B" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!" & T_FreqStartRng & "," & SheetName & "$" & T_FreqStartRng & ":$M$6,0))" '<----note that SheetName includes apostrophe character and ActiveSheet.Name does not.....trickyyyyy
        ExtendFunction
        
        ElseIf Left(T_SheetType, 2) = "TO" Then
        Cells(Selection.Row, DestinationCol).Value = "=INDEX(" & SheetName & "$E$" & FirstRow & ":$Y$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
        "," & SheetName & "$B$" & FirstRow & ":$B" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!" & T_FreqStartRng & "," & SheetName & "$" & T_FreqStartRng & ":$Y$6,0))"
        ExtendFunction
        
        ElseIf T_SheetType = "MECH" Then
        Cells(Selection.Row, DestinationCol).Value = "=INDEX(" & SheetName & "$B$" & FirstRow & ":$M$" & LastRow & ",MATCH('" & ActiveSheet.Name & "'!$B$" & Selection.Row & _
        "," & SheetName & "$B$" & FirstRow & ":$B" & LastRow & ",0),MATCH('" & ActiveSheet.Name & "'!T$6," & SheetName & "$B$6:$M$6,0))"
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
' Name:     ToggleActiveSelection
' Author:   AA
' Desc:     Toggles a cell inactive or active. Same as Trace version but can do
'           multiple selections or rows and doesn't require a Typecode.
' Args:     Just the selected cells
' Comments: (1)
'==============================================================================

Sub ToggleActiveSelection()
Dim startCol As Integer, endCol As Integer, WrkRow As Integer, WrkArea As Integer
Dim startRow As Integer, endRow As Integer
Dim CharKeep As Integer
Dim Orig As String, OrigFmt As String, FormatArchive As String
Dim NewValue As String

Application.ScreenUpdating = False

Set SelectedCells = Selection

Set WrkRng = SelectedCells.Areas(1)
startRow = WrkRng.Rows(1).Row
endRow = startRow + WrkRng.Rows.Count - 1
startCol = WrkRng.Columns(1).Column
endCol = startCol + WrkRng.Columns.Count - 1

' If currently inactive, toggle active. Looks for [,"\*] within formula.
If InStr(1, Cells(startRow, startCol).Formula, ",""\*") <> 0 Then
    
    For WrkArea = 1 To SelectedCells.Areas.Count
        Set WrkRng = SelectedCells.Areas(WrkArea)
        startRow = WrkRng.Rows(1).Row
        endRow = startRow + WrkRng.Rows.Count - 1
        startCol = WrkRng.Columns(1).Column
        endCol = startCol + WrkRng.Columns.Count - 1
    
        For WrkRow = startRow To endRow
            For i = startCol To endCol
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
                    'do nothing
                End If
            Next i
        Next WrkRow
    Next WrkArea
    
ElseIf IsEmpty(Cells(startRow, startCol)) = False Then
    
    For WrkArea = 1 To SelectedCells.Areas.Count
        Set WrkRng = SelectedCells.Areas(WrkArea)
        startRow = WrkRng.Rows(1).Row
        endRow = startRow + WrkRng.Rows.Count - 1
        startCol = WrkRng.Columns(1).Column
        endCol = startCol + WrkRng.Columns.Count - 1
        
        For WrkRow = startRow To endRow
            For i = startCol To endCol
                ' If currently inactive, toggle active. Looks for [,"\*] within formula.
                If InStr(1, Cells(WrkRow, i).Formula, ",""\*") <> 0 Then
                    'do nothing
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
        Next WrkRow
    Next WrkArea
End If

Application.ScreenUpdating = True

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

SetDescription "Correction"
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
    SetDescription "Total"
    Else
    SetDescription LineDescStr
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
    If ApplyStyleCode = "AutoSum_Total" Then
    SetTraceStyle "Total"
    ElseIf ApplyStyleCode = "AutoSum_Subtotal" Then
    SetTraceStyle "Subtotal"
    ElseIf ApplyStyleCode = "AutoSum_Normal" Then
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
' Comments: (1)updated to simplify method, only implementation rather than
'           branching for different sheet types
'==============================================================================
Sub OneThirdsToOctave()

Dim splitAddr() As String 'array for extracting elements from range
Dim SheetName As String 'for referencing the name of the sheet
Dim WriteRw As Integer 'first row where the result is going
Dim Rw As Integer 'row inside loop
Dim Col As Integer 'column inside loop
Dim RwStart As Integer 'first row to be coverted
Dim RwEnd As Integer 'last row to be converted
Dim ColStart As Integer 'first column result
Dim ColEnd As Integer 'last columnt result
Dim refCol As Integer 'input column, skips by 3 within loop
Dim targetRange As String 'range string to feed into function, built in the loop

    'check for sheet types
    If Left(T_SheetType, 3) <> "OCT" And T_SheetType <> "CVT" Then
    ErrorTypeCode
    End If
    
    'set form controls
    If T_SheetType = "CVT" Then
    frmConvert.refRangeSelector.Enabled = False
    frmConvert.refRangeSelector.Value = Selection.Address
    Else
    frmConvert.refRangeSelector.Enabled = True
    End If

'call the form
frmConvert.Show

    'catch errors from frmConvert
    If btnOkPressed = False Then End
    If UserSelectedAddress = "" Then End

splitAddr = Split(UserSelectedAddress, "$", Len(UserSelectedAddress), vbTextCompare)

SheetName = splitAddr(LBound(splitAddr)) 'sheet is the first element

'set initial row
WriteRw = Selection.Row

    'set range of rows
    If UBound(splitAddr) >= 3 Then 'range of cells, not just a single cell
    RwStart = CInt(Left(splitAddr(2), Len(splitAddr(2)) - 1))
    RwEnd = CInt(splitAddr(4))
    Else
    RwStart = CInt(splitAddr(2))
    RwEnd = RwStart
    End If
    
    'loop through each row
    For Rw = RwStart To RwEnd
    
        'reset ColStart and CoEnd to start and of bands
        If T_SheetType = "OCT" Then
        ColStart = T_RegenStart + 1 'start one band over becase TO starts from 50Hz band
        ColEnd = T_RegenEnd - 1 'finish one band early for the same reason
        Else
        ColStart = T_RegenStart
        ColEnd = T_RegenEnd
        End If
        
    'reset refCol
    refCol = T_LossGainStart
    
    SetDescription "Conversion from one-thirds - " & SumAverageMode, WriteRw
    
        'loop through each column
        For Col = ColStart To ColEnd
        
        targetRange = Range(Cells(Rw, refCol), _
            Cells(Rw, refCol + 2)).Address(False, False)
            
            'build formula based on the mode
            Select Case SumAverageMode 'selected from radio boxes in form frmConvert
            Case Is = "Sum"
            Cells(WriteRw, Col).Value = "=SPLSUM(" & SheetName & targetRange & ")"
            Case Is = "Average"
            Cells(WriteRw, Col).Value = "=AVERAGE(" & SheetName & targetRange & ")"
            Case Is = "Log Av"
            Cells(WriteRw, Col).Value = "=SPLAV(" & SheetName & targetRange & ")"
            Case Is = "TL"
            Cells(WriteRw, Col).Value = "=TL_ThirdsToOctave(" & SheetName & targetRange & ")"
            End Select
            
        refCol = refCol + 3
        
        Next Col
    WriteRw = WriteRw + 1
    Next Rw
    
'apply styles
If T_SheetType <> "CVT" Then SetTraceStyle "Reference"

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
        Cells(Selection.Row, T_Description) = _
            Cells(Selection.Row - 1, T_Description).Value & " (Linear)"
        Else
        Cells(Selection.Row, T_Description) = _
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
' Name:     BuildFormula
' Author:   PS
' Desc:     Creates formula in chosen range
' Args:     FormulaStr - Formula to be included
'           IsRegen - set to true for Regen columns
' Comments: (1)
'==============================================================================
Sub BuildFormula(FormulaStr, Optional IsRegen As Boolean)
    If IsEmpty(IsRegen) Then IsRegen = False
    
'Debug.Print FormulaStr
    
    If IsRegen = True Then
    Cells(Selection.Row, T_RegenStart).Formula = "=" & FormulaStr
    Else 'default to LossGain
    Cells(Selection.Row, T_LossGainStart).Formula = "=" & FormulaStr
    End If
ExtendFunction (IsRegen)
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
    Range(Cells(Selection.Row, T_RegenStart), _
        Cells(Selection.Row, T_RegenEnd)).PasteSpecial (xlPasteFormulas)
    Else
    Cells(Selection.Row, T_LossGainStart).Copy
    Range(Cells(Selection.Row, T_LossGainStart), _
        Cells(Selection.Row, T_LossGainEnd)).PasteSpecial (xlPasteFormulas)
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
Sub ParameterMerge(Rw As Integer, Optional NumCols As Integer)
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
Sub ParameterUnmerge(Rw As Integer)

Range(Cells(Rw, T_ParamStart), Cells(Rw, T_ParamEnd)).UnMerge

    With Range(Cells(Rw, T_ParamStart), Cells(Rw, T_ParamEnd))
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
Sub CreateSparkline(Rw As Integer, p_index As Integer, Optional colorIndex As Integer)

Dim DataRangeStr As String

    If p_index <= UBound(T_ParamRng) Then
        'Set where the data to be graphed is
        If T_SheetType = "MECH" Then
        DataRangeStr = Range(Cells(Rw, T_RegenStart), _
            Cells(Rw, T_RegenEnd)).Address
        Else
        DataRangeStr = Range(Cells(Rw, T_LossGainStart), _
            Cells(Rw, T_LossGainEnd)).Address
        End If
        
    'add the sparkline
    Cells(Rw, T_ParamStart + p_index).SparklineGroups.Add _
        Type:=xlSparkLine, SourceData:=DataRangeStr
        
        'formatting etc
        With Cells(Rw, T_ParamStart + p_index).SparklineGroups.Item(1)
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
            For Col = 2 To 7
            Sheets("SUMMARY").Cells(WriteRw, Col + 5).Value = _
                "='" & Sheets(sh).Name & "'!" & Cells(29, Col + 5).Address
            Sheets("SUMMARY").Cells(WriteRw, Col + 5).NumberFormat = "0.0"
            Next Col
            
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

'==============================================================================
' Name:     SetDescription
' Author:   PS
' Desc:     Puts in description, only if there's nothing in there, otherwise
'           puts the name in the comment
' Args:     None
' Comments: (1)
'==============================================================================
Sub SetDescription(DescriptionString As String, Optional InputRw As Integer)
Dim Rw As Integer

    'set row
    If InputRw = 0 Or IsMissing(InputRw) Then
    Rw = Selection.Row
    Else
    Rw = InputRw
    End If

    'check for description field
    If Cells(Rw, T_Description).Value = "" Then 'set description
    Cells(Rw, T_Description).ClearContents
    Cells(Rw, T_Description).ClearComments
    Cells(Rw, T_Description).Value = DescriptionString
    Else 'add as comment
    InsertComment DescriptionString, T_Description, True, Rw
    End If
    
End Sub

'==============================================================================
' Name:     SetDataValidation
' Author:   PS
' Desc:     Sets data validation for cells for given column number
' Args:     col - column number
'           ValidationOptionsStr - String of options to be set
' Comments: (1)
'==============================================================================
Sub SetDataValidation(Col As Integer, ValidationOptionStr As String)
    With Cells(Selection.Row, Col).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:=ValidationOptionStr
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
    End With
End Sub


'==============================================================================
' Name:     InsertComment
' Author:   PS
' Desc:     Inserts comment into a given column number
' Args:     CommentStr - String of comment to be added
'           col - column number
'           append - set to TRUE to append the comment to the existing comment
' Comments: (1) Deletes previous comment unless append is TRUE
'==============================================================================
Sub InsertComment(CommentStr As String, Col As Integer, _
    Optional append As Boolean, Optional InputRw As Integer)
    
Dim CheckRng As Range
Dim Rw As Integer

    'set row
    If InputRw = 0 Or IsMissing(InputRw) Then
    Rw = Selection.Row
    Else
    Rw = InputRw
    End If
    
'add comment with more detail
Set CheckRng = Cells(Rw, Col)

    If Not CheckRng.Comment Is Nothing Then
        If append = False Then 'delete!
        CheckRng.Comment.Delete
        Else 'append, then clear what's in there
        CommentStr = CheckRng.Comment.Text & " // " & CommentStr
        CheckRng.Comment.Delete
        End If
    End If

CheckRng.AddComment CommentStr
CheckRng.Comment.Shape.TextFrame.AutoSize = True
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

