Attribute VB_Name = "RowOperations"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
Public UserSelectedAddress As String
Public UserDestinationAddress As String
Public SumAverageMode As String
Public LookupMultiRow As Boolean
Public DynamicReferencing As Boolean
Public AddSchedMarker As Boolean
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
Function GetSheetName(inputStr As String) 'Sheet name, first row, last row
Dim SplitStr() As String
SplitStr = Split(inputStr, "!", Len(inputStr), vbTextCompare)
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
Function GetFirstRow(inputStr As String)
Dim SplitStr() As String
SplitStr = Split(inputStr, "$", Len(inputStr), vbTextCompare)
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
Function GetLastRow(inputStr As String)
Dim SplitStr() As String
SplitStr = Split(inputStr, "$", Len(inputStr), vbTextCompare)
GetLastRow = CInt(SplitStr(UBound(SplitStr)))
End Function

'==============================================================================
' Name:     HasDataValidation
' Author:   PS
' Desc:     Returns TRUE if the range has validation on it
' Args:     rng - the range to be tested
' Comments: (1)
'==============================================================================
Function HasDataValidation(rng As Range) As Boolean

On Error Resume Next

DVtype = rng.Validation.Type

On Error GoTo 0

    If DVtype = 3 Then
    HasDataValidation = True
    Else
    HasDataValidation = False
    End If

End Function

'==============================================================================
' Name:     GenerateAddress
' Author:   PS
' Desc:     Returns an address, formatted correctly, with commas and references
' Args:
' Comments: (1) TODO: Finish this function
'==============================================================================
Function GenerateAddress(SheetName As String, rw As Long, Col As Long, _
    Optional AbsoluteRow As Boolean = True, Optional AbsoluteCol As Boolean = True, _
    Optional NumRows As Long = 1, Optional NumCols As Long = 1)
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(SheetName)

GenerateAddress = ws.Range(ws.Cells(rw, Col).Resize(NumRows, NumCols).Address(AbsoluteRow, AbsoluteCol))

End Function

'==============================================================================
' Name:     FindTopOfBlock
' Author:   PS
' Desc:     Finds the top of the calculation block by looking for the next
'           blank row above the selected line
' Args:     Col - the Column to be analysed
'           startRw - option to start at any row, not just the current one
' Comments: (1)
'==============================================================================
Function FindTopOfBlock(Col As Integer, Optional startRw As Integer)

Dim FoundRw As Boolean
Dim rw As Integer

If startRw = 0 Or IsMissing(startRw) Then
    rw = Selection.Row
Else
    rw = startRw
End If

'loop to find end start of calculation
While FoundRw = False
    rw = rw - 1
    
    If rw < 8 Then 'A weighting is on line 7 for all template sheets
        rw = 7 'A weighting line is the same as a blank line
        FoundRw = True
    ElseIf Cells(rw, Col).Value = "" Then
        FoundRw = True
    End If
        
Wend

FindTopOfBlock = rw + 1 'it's the next one down, not the blank one

End Function

'<-TODO this function
''==============================================================================
'' Name:     CalcZoneType
'' Author:   PS
'' Desc:     Returns TRUE if the range has validation on it
'' Args:     rng - the range to be tested
'' Comments: (1)
''==============================================================================
'Function CalcZoneType(rng As Range) As Integer
'
'On Error Resume Next
'
'DVtype = rng.Validation.Type
'
'On Error GoTo 0
'
'    If DVtype = 3 Then
'    HasDataValidation = True
'    Else
'    HasDataValidation = False
'    End If
'
'End Function



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
Dim LastRow As Integer
    
    If SkipUserInput = False Then 'runs by default
        If Selection.Rows.Count > 1 Then
        ClearMultipleRows = MsgBox("Are you sure you want to clear selected rows?", _
            vbYesNo, "Check, one, two...")
        End If
        
        If ClearMultipleRows = vbNo Then End 'catch user abort
    End If

Application.ScreenUpdating = False
LastRow = Selection.Row + Selection.Rows.Count - 1

    'loop for clearing data, works
    For rw = Selection.Row To LastRow
    Application.StatusBar = "Clearing row " & rw & "/" & LastRow
    'Description
    Cells(rw, 1).ClearContents
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
    
Application.StatusBar = False
Application.ScreenUpdating = True

End Sub

'==============================================================================
' Name:     FlipSign
' Author:   PS
' Desc:     Makes selected calls (or row) negative values. Maintains the
'           formula and will revert back.
' Args:     None
' Comments: (1) Updated to skip empty cells
'==============================================================================
Sub FlipSign()
Dim rw As Integer
Dim startCol As Integer
Dim endCol As Integer

    For rw = Selection.Row To Selection.Row + Selection.Rows.Count - 1
        For Col = T_LossGainStart To T_LossGainEnd
        
            'two cases, depending on formulas or static values
            If Cells(rw, Col).HasFormula Then
                'check if second character of formula is minus
                If Mid(Cells(rw, Col).Formula, 2, 1) = "-" Then
                Cells(rw, Col).Formula = Replace(Cells(rw, Col).Formula, _
                    "=-", "=", 1, Len(Cells(rw, Col).Formula), vbTextCompare)
                Else
                Cells(rw, Col).Formula = Replace(Cells(rw, Col).Formula, _
                    "=", "=-", 1, Len(Cells(rw, Col).Formula), vbTextCompare)
                End If
            Else 'static value
                'check if cell is empty
                If IsEmpty(Cells(rw, Col)) = False Then
                Cells(rw, Col).Value = Cells(rw, Col).Value * -1
                End If
            End If
            
        Next Col
    Next rw

End Sub

'==============================================================================
' Name:     MoveUp
' Author:   PS
' Desc:     Moves selected calculation row(s) up by 1.
' Args:     None
' Comments: (1)
'==============================================================================
Sub MoveUp()
Dim startRw As Integer
Dim endRw As Integer
Dim LastCol As Integer

Application.ScreenUpdating = False

startRw = Selection.Row
endRw = Selection.Row + Selection.Rows.Count - 1

    'check for adjacent empty row
    If Cells(startRw - 1, 2).Value <> "" Then
    msg = MsgBox("There appears to be data in the cells above.  " & _
        "Continuing will delete this data. Do you want to continue?", _
        vbYesNo, "Check yo'self")
        If msg = vbNo Then End
    End If

'TODO: move in three parts to preserve merged cells?
LastCol = Application.WorksheetFunction.Max(T_LossGainEnd, T_ParamEnd)
'formulas
Range(Cells(startRw, T_Description), Cells(endRw, LastCol)).Cut _
    Destination:=Range(Cells(startRw - 1, T_Description), Cells(endRw - 1, LastCol))
'formats
Range(Cells(startRw - 1, T_Description), Cells(startRw - 1, LastCol)).Copy
Range(Cells(endRw, T_Description), Cells(endRw, LastCol)).PasteSpecial Paste:=xlPasteFormats

ClearRow (True)
    
'move to select lower row
Range(Cells(startRw - 1, T_Description), Cells(endRw - 1, T_Description)).Select

Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub

'==============================================================================
' Name:     MoveDown
' Author:   PS
' Desc:     Moves selected calculation row(s) down by 1.
' Args:     None
' Comments: (1)
'==============================================================================

Sub MoveDown()
Dim startRw As Integer
Dim endRw As Integer

Application.ScreenUpdating = False

startRw = Selection.Row
endRw = Selection.Row + Selection.Rows.Count - 1

    'check for adjacent empty row
    If Cells(endRw + 1, 2).Value <> "" Then
    msg = MsgBox("There appears to be data in the cells below.  " & _
        "Continuing will delete this data. Do you want to continue?", _
        vbYesNo, "Check yo'self")
        If msg = vbNo Then End
    End If

'Range("B" & startRw & ":D" & endRw).Cut Destination:=Range("B" & startRw + 1 & ":D" & endRw + 1) 'Description
'    'TODO: update for Trace 3
'    If Left(T_SheetType, 3) = "OCT" Then
'    Range("E" & startRw & ":O" & endRw).Cut Destination:=Range("E" & startRw + 1 & ":O" & endRw + 1) 'Formulas
'    Range("B" & startRw + 1 & ":O" & startRw + 1).Copy 'formats
'    Range("B" & startRw & ":O" & startRw).PasteSpecial Paste:=xlPasteFormats
'    ElseIf Left(T_SheetType, 2) = "TO" Then
'    Range("E" & startRw & ":AA" & endRw).Cut Destination:=Range("E" & startRw + 1 & ":AA" & endRw + 1) 'Formulas
'    Range("B" & startRw + 1 & ":AA" & startRw + 1).Copy 'formats
'    Range("B" & startRw & ":AA" & startRw).PasteSpecial Paste:=xlPasteFormats
'    ElseIf T_SheetType = "LF_TO" Then
'    Range("E" & startRw & ":AG" & endRw).Cut Destination:=Range("E" & startRw + 1 & ":AG" & endRw + 1) 'Formulas
'    Range("B" & startRw + 1 & ":AG" & startRw + 1).Copy 'formats
'    Range("B" & startRw & ":AG" & startRw).PasteSpecial Paste:=xlPasteFormats
'    End If
'
''move to select lower row
'Range(Cells(startRw + 1, 2), Cells(endRw + 1, 2)).Select

'TODO: move in three parts to preserve merged cells?
LastCol = Application.WorksheetFunction.Max(T_LossGainEnd, T_ParamEnd)
'formulas
Range(Cells(startRw, T_Description), Cells(endRw, LastCol)).Cut _
    Destination:=Range(Cells(startRw + 1, T_Description), Cells(endRw + 1, LastCol))
'formats
Range(Cells(startRw + 1, T_Description), Cells(startRw + 1, LastCol)).Copy
Range(Cells(startRw, T_Description), Cells(startRw, LastCol)).PasteSpecial Paste:=xlPasteFormats

ClearRow (True)
    
'move to select lower row
Range(Cells(startRw + 1, T_Description), Cells(endRw + 1, T_Description)).Select

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
Dim startRw As Integer
Dim endRw As Integer
Dim startCol As Integer
Dim endCol As Integer

Application.ScreenUpdating = False

startRw = Selection.Row
endRw = Selection.Row + Selection.Rows.Count - 1
'startCol = GetSheetTypeColumns(SheetType, "LossGainStart")
'endCol = GetSheetTypeColumns(SheetType, "LossGainEnd")

Range(Cells(startRw, T_LossGainStart + 1), Cells(endRw, T_LossGainEnd)).Copy
Cells(startRw, T_LossGainStart).PasteSpecial Paste:=xlPasteValues
Cells(startRw, T_LossGainEnd).ClearContents

Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub

 '==============================================================================
' Name:     MoveRight
' Author:   PS
' Desc:     Moves selected spectrum one band to the right
' Args:     None
' Comments: (1) Applies to main columns only, descriptions stay put
'===============================================================================
Sub MoveRight()
Dim startRw As Integer
Dim endRw As Integer
Dim startCol As Integer
Dim endCol As Integer

Application.ScreenUpdating = False

startRw = Selection.Row
endRw = Selection.Row + Selection.Rows.Count - 1
'startCol = GetSheetTypeColumns(SheetType, "LossGainStart")
'endCol = GetSheetTypeColumns(SheetType, "LossGainEnd")

Range(Cells(startRw, T_LossGainStart), Cells(endRw, T_LossGainEnd - 1)).Copy
Cells(startRw, T_LossGainStart + 1).PasteSpecial Paste:=xlPasteValues
Cells(endRw, T_LossGainStart).ClearContents

Application.CutCopyMode = False

Application.ScreenUpdating = True

End Sub

 '==============================================================================
' Name:     Row Reference
' Author:   PS
' Desc:     Puts in a reference to data in another row or sheet
' Args:     None
' Comments: (1) Updated to work across sheets
'           (2) Updated to 'push' or 'pull' referenced range
'           (3) Selects SWL as the default on row 8 of MECH noise sheets
'===============================================================================
Sub RowReference()
Dim FirstRow, LastRow As Integer
Dim LookupSheetName, DestinationSheetName  As String
Dim DestinationCol As Integer ' for checking different layouts
Dim DestinationRw As Integer
Dim ReferenceMethod As String

'clear old values
frmRowReference.refDestinationSelector.Value = ""
frmRowReference.refRangeSelector.Value = ""
    
'Allow regen for MECH sheet layout, and enable controls
If T_SheetType = "MECH" Then
    frmRowReference.optRegenSWL.Enabled = True
        'default to SWL on the top of the sheet
        If Selection.Row = 8 Then frmRowReference.optRegenSWL.Value = True
Else
    frmRowReference.optRegenSWL.Enabled = False
End If


    'Set default values, based on method
If Selection.Rows.Count > 1 Then  'put in the loopup range from selected
    frmRowReference.refRangeSelector.Value = _
        "'" & ActiveSheet.Name & "'!" & Range(Selection.Address).Address
Else 'one line selected, pull reference in from elsewhere
       
    If HasDataValidation(Cells(Selection.Row, T_Description)) Then 'read existing validation and put in the form
        frmRowReference.refRangeSelector.Value = _
            Cells(Selection.Row, T_Description).Validation.Formula1
    Else
        frmRowReference.refRangeSelector.Value = ""
    End If
        
frmRowReference.refDestinationSelector.Value = _
    "'" & ActiveSheet.Name & "'!" & Selection.Address
End If

'show the form for inputs
frmRowReference.Show

    'error catch
    If btnOkPressed = False Then End
    If UserSelectedAddress = "" Then End
    
LookupSheetName = GetSheetName(UserSelectedAddress)
DestinationSheetName = GetSheetName(UserDestinationAddress)
FirstRow = GetFirstRow(UserSelectedAddress)
LastRow = GetLastRow(UserSelectedAddress)

'set destination column
If RegenDestinationRange = True Then
    DestinationCol = T_RegenStart
Else
    DestinationCol = T_LossGainStart
End If
    
DestinationRw = GetFirstRow(UserDestinationAddress)
Cells(DestinationRw, T_Description).Select
    'TODO: check for mismatched sheet types
    'something like:
'    If Sheets(LookupSheetName).Range("TYPECODE") <> T_SheetType Then
'
'    End If

If LookupMultiRow = False Then
    If DynamicReferencing = True Then
        SetDescription "=CONCAT(""Ref: ""," & LookupSheetName & "$B$" & FirstRow & ")"
    Else
        SetDescription Cells(FirstRow, T_Description).Value
    End If
    
    If AddSchedMarker = True Then
    ApplyTraceMarker ("Schedule")
    End If
    
    Cells(DestinationRw, DestinationCol).Value = "=" & _
        LookupSheetName & "E$" & FirstRow
    ExtendFunction (RegenDestinationRange)
    
Else 'multimode = true
    
    'trim equals character if already there
    If Left(LookupSheetName, 1) = "=" Then
        LookupSheetName = Right(LookupSheetName, Len(LookupSheetName) - 1)
    End If
        
    SetDataValidation T_Description, "=" & LookupSheetName & "$B$" & FirstRow & _
        ":$B$" & LastRow, DestinationRw
        
    'select first entry by default
    SetDescription Range(LookupSheetName & "$B$" & FirstRow)
    
        'create index-match formula <--TODO: update this for Trace 3
    If T_SheetType = "OCT" Then
        Cells(DestinationRw, DestinationCol).Value = _
            "=INDEX(" & LookupSheetName & "$E$" & FirstRow & ":$M$" & LastRow & _
            ",MATCH('" & ActiveSheet.Name & "'!$B" & DestinationRw & "," & LookupSheetName & "$B$" & FirstRow & ":$B$" & LastRow & ",0)," & _
            "MATCH('" & ActiveSheet.Name & "'!" & T_FreqStartRng & "," & LookupSheetName & "$" & T_FreqStartRng & ":$M$6,0))" '<----note that SheetName includes apostrophe character and ActiveSheet.Name does not.....trickyyyyy
        ExtendFunction
        
    ElseIf Left(T_SheetType, 2) = "TO" Then
        Cells(DestinationRw, DestinationCol).Value = _
            "=INDEX(" & LookupSheetName & "$E$" & FirstRow & ":$Y$" & LastRow & _
            ",MATCH('" & ActiveSheet.Name & "'!$B" & DestinationRw & "," & LookupSheetName & "$B$" & FirstRow & ":$B$" & LastRow & ",0)," & _
            "MATCH('" & ActiveSheet.Name & "'!" & T_FreqStartRng & "," & LookupSheetName & "$" & T_FreqStartRng & ":$Y$6,0))"
        ExtendFunction
        
    ElseIf T_SheetType = "MECH" Then
        Cells(DestinationRw, DestinationCol).Value = _
            "=INDEX(" & LookupSheetName & "$B$" & FirstRow & ":$M$" & LastRow & _
            ",MATCH('" & ActiveSheet.Name & "'!$B" & DestinationRw & "," & LookupSheetName & "$B$" & FirstRow & ":$B$" & LastRow & ",0)," & _
            "MATCH('" & ActiveSheet.Name & "'!T$6," & LookupSheetName & "$B$6:$M$6,0))"
        ExtendFunction (RegenDestinationRange)
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
' Comments: (1) Edited to make column selection dependant on cursor position
'           (2) Updated to work on multiple rows
'           (3) Turn off calculation while this happens
'           (4) Turn off screen updating
'==============================================================================
Sub ToggleActive()
Dim startCol As Integer, endCol As Integer
Dim startRw As Integer, endRw As Integer, WrkRow As Integer
Dim CharKeep As Integer
Dim Orig As String, OrigFmt As String, FormatArchive As String
Dim NewValue As String

Application.Calculation = xlManual
Application.ScreenUpdating = False

WrkRow = Selection.Row

    'set working range based on where the cursor is
    If Selection.Column >= T_LossGainStart And Selection.Column <= T_LossGainEnd Then
    startCol = T_LossGainStart
    endCol = T_LossGainEnd
    ElseIf Selection.Column >= T_RegenStart And Selection.Column <= T_RegenEnd Then
    startCol = T_RegenStart
    endCol = T_RegenEnd
    Else 'fallback
    startCol = T_LossGainStart
    endCol = T_LossGainEnd
    End If
    
startRw = Selection.Row
endRw = Selection.Row + Selection.Rows.Count - 1
'CHECK FOR NON HEADER ROWS
CheckTemplateRow (startRw)

    For WrkRow = startRw To endRw
    Application.StatusBar = "Toggle Row: " & WrkRow
        ' Toggles between active and inactive
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

Application.Calculation = xlAutomatic
Application.StatusBar = False
Application.ScreenUpdating = True

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
Sub SingleCorrection(Optional DefaultValue As Integer)

If IsMissing(DefaultValue) Then DefaultValue = -5

SetDescription "Correction"
BuildFormula "" & T_ParamRng(0)
Cells(Selection.Row, T_ParamStart).Value = DefaultValue
ParameterMerge (Selection.Row)
SetUnits "dB", T_ParamStart, 0

SetTraceStyle "Input", True
Cells(Selection.Row, T_ParamStart).Select 'move to parameter column t set value
End Sub

'==============================================================================
' Name:     AutoSum_UserInput
' Author:   PS
' Desc:     Catches the button call on the ribbon for user selection
' Args:
' Comments: (1)
'==============================================================================
Sub AutoSum_UserInput(Optional ApplyStyleCode As String, Optional LineDescStr As String)
AutoSum ApplyStyleCode, LineDescStr, True
End Sub


'==============================================================================
' Name:     AutoSum
' Author:   PS
' Desc:     Sums all rows until a blank row is reached.
' Args:     ApplyStyleCode - String that says what style to use
'           LineDescStr - String for the T_description column
' Comments: (1) TODO: needs update to scan first active column, if using
'            a reduced working range
'==============================================================================
Sub AutoSum(Optional ApplyStyleCode As String, Optional LineDescStr As String, _
    Optional AskUserInput As Boolean)
Dim FindRw As Integer
Dim ScanCol As Integer
Dim FoundRw As Boolean
Dim NumRows As Integer

'get number of rows above, and put it in the row selector form
T_FirstSelectedRow = Selection.Row
FindRw = FindTopOfBlock(Selection.Column)
If AskUserInput = True Then
    frmRowSelector.sbRowsAbove.Value = Selection.Row - FindRw
    frmRowSelector.Show
        If btnOkPressed = False Then End
    'move down for summation
    Cells(T_LastSelectedRow + 1, T_LossGainStart).Select
Else
    T_FirstSelectedRow = FindRw
    T_LastSelectedRow = Selection.Row - 1
End If

If LineDescStr = "" Then
    SetDescription "Total"
Else
    SetDescription LineDescStr
End If

'build formula, special case for sound power
If ApplyStyleCode = "AutoSum_Source" Then
    BuildFormula "SPLSUM(" & _
        Range(Cells(T_FirstSelectedRow, T_LossGainStart), _
        Cells(T_LastSelectedRow, T_LossGainStart)).Address(False, False) & ")"
Else 'all others are normal sum
    BuildFormula "SUM(" & _
        Range(Cells(T_FirstSelectedRow, T_LossGainStart), _
        Cells(T_LastSelectedRow, T_LossGainStart)).Address(False, False) & ")"
End If

'Limit the options to the three main styles
If ApplyStyleCode = "AutoSum_Total" Then
    SetTraceStyle "Total"
ElseIf ApplyStyleCode = "AutoSum_Subtotal" Then
    SetTraceStyle "Subtotal"
ElseIf ApplyStyleCode = "AutoSum_Normal" Then
    SetTraceStyle "Normal"
ElseIf ApplyStyleCode = "AutoSum_Source" Then
    SetTraceStyle "Lw Source"
Else 'default to Subtotal
    SetTraceStyle "Subtotal"
End If

ApplyTraceMarker ("Sum")

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
Dim rw As Integer 'row inside loop
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
    For rw = RwStart To RwEnd
    
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
        
        targetRange = Range(Cells(rw, refCol), _
            Cells(rw, refCol + 2)).Address(False, False)
            
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
    Next rw
    
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
BuildFormula "" & _
    Cells(Selection.Row - 1, T_LossGainStart).Address(False, False) & "+" & _
    Cells(7, T_LossGainStart).Address(True, False)

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
    If NamedRangeExists("TYPECODE") Then 'send to normal conversion form
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
' Comments: (1) Updated to catch equals sign - can have it or not, either way.
'           (2) TODO: include setting description text as part of this method
'==============================================================================
Sub BuildFormula(FormulaStr, Optional IsRegen As Boolean)

Dim FirstCharacter As String

    If IsEmpty(IsRegen) Then IsRegen = False
    
'Debug.Print FormulaStr

'catch missing equals sign
If Left(FormulaStr, 1) = "=" Then
    FirstCharacter = "" 'no need to add the equals
Else
    FirstCharacter = "=" 'add the equals sign
End If

'Write the formula to the correct area
If IsRegen = True Then
    Cells(Selection.Row, T_RegenStart).Formula = FirstCharacter & FormulaStr
Else 'default to LossGain
    Cells(Selection.Row, T_LossGainStart).Formula = FirstCharacter & FormulaStr
End If
    
ExtendFunction (IsRegen)

End Sub

'==============================================================================
' Name:     ExtendFunction
' Author:   PS
' Desc:     Copies formulas to the correct ranges
' Args:     None
' Comments: (1) Sheet currently supports OCT, OCTA, TO, TOA sheets
'           (2) updated to loop through all rows
'           (3) updated to extend from the first column with value onwards
'==============================================================================
Sub ExtendFunction(Optional ApplyToRegen As Boolean)
Dim StartAddr As String
Dim rw As Integer
Dim Col As Integer
Dim CopyColStart, CopyRowStart As Integer 'the copying will start here
Dim CopyColEnd, CopyRowEnd As Integer 'the copying will end here
Dim FormulaFound As Boolean


'note starting position
StartAddr = Selection.Address
Application.ScreenUpdating = False

    If T_LossGainStart < 1 Or T_LossGainEnd < 1 Then
    'public variables not defined, let's fix that
    SetSheetTypeControls
    End If
    
CopyRowStart = Selection.Row
CopyRowEnd = Selection.Row + Selection.Rows.Count - 1

'loop for each row
For rw = CopyRowStart To CopyRowEnd
    
    'set copying range
    If ApplyToRegen = True Then
        CopyColStart = T_RegenStart
        CopyColEnd = T_RegenEnd
    Else
        CopyColStart = T_LossGainStart
        CopyColEnd = T_LossGainEnd
    End If
    
    'loop through columns to find first formula
    For Col = CopyColStart To CopyColEnd
        
        'find cell to copy
        If FormulaFound = False And IsEmpty(Cells(rw, Col)) = False Then
            Cells(rw, Col).Copy
            FormulaFound = True 'change flag
        End If
        
    Next Col
    
    'loop through columns to paste the formulas
    For Col = CopyColStart To CopyColEnd
            'copy across, skipping excluded frequency columns
        If Right(Cells(T_FreqRow, Col).Value, 1) <> "*" And FormulaFound = True Then
            Range(Cells(rw, Col), Cells(rw, Col)).PasteSpecial (xlPasteFormulas)
        End If
    Next Col
    
    FormulaFound = False 'reset this flag for the next row
Next rw
    
'remove trailing value if the first column is disabled
If Right(Cells(T_FreqRow, CopyColStart).Value, 1) = "*" Then
    Range(Cells(CopyRowStart, CopyColStart), _
        Cells(CopyRowEnd, CopyColStart)).ClearContents
End If
    
Application.CutCopyMode = False
'go back to starting position
Range(StartAddr).Select
Application.ScreenUpdating = True
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
'           (2) Used variable TargetRng and checks for merged property. Neater.
'           (3) Loops through all param cells to unmerge
'==============================================================================
Sub ParameterUnmerge(rw As Integer)

Dim TargetRng As Range
Dim Col As Integer

'old version
'Set TargetRng = Range(Cells(rw, T_ParamStart), Cells(rw, T_ParamEnd))

For Col = T_ParamStart To T_ParamEnd
Set TargetRng = Range(Cells(rw, Col), Cells(rw, Col))
    If TargetRng.MergeCells Then 'check for merged property
    TargetRng.UnMerge
    End If
Next Col

'borders
Set TargetRng = Range(Cells(rw, T_ParamStart), Cells(rw, T_ParamEnd))
    With TargetRng
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
Sub SelectNextRow(Optional Col As Integer)

If Col = 0 Then
    Cells(Selection.Row + 1, Selection.Column).Select 'move down
Else
    Cells(Selection.Row + 1, Col).Select 'move down
End If

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
        If NamedRangeExists("TYPECODE") Then
            If Range("TYPECODE").Value = "RT" Then
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
        End If
    Next sh
Sheets("SUMMARY").Activate
End Sub

'==============================================================================
' Name:     SetDescription
' Author:   PS
' Desc:     Puts in description, only if there's nothing in there, otherwise
'           puts the name in the comment
' Args:     DescriptionString - Text to be inserted
'           InputRw - Row number of description, defaults to selection.row
'           OverWriteExisting - set to TRUE to replace exisitng text
' Comments: (1) Updated to allow force overwrite
'==============================================================================
Sub SetDescription(DescriptionString As String, Optional InputRw As Integer, _
Optional OverWriteExisting As Boolean)

Dim rw As Integer

    'set row
    If InputRw = 0 Or IsMissing(InputRw) Then
    rw = Selection.Row
    Else
    rw = InputRw
    End If
    
    'check for the same description being there already
    If Cells(rw, T_Description).Formula = DescriptionString Then Exit Sub
    
    'check for description field
    If Cells(rw, T_Description).Value = "" Or _
        OverWriteExisting = True Then 'set description
    Cells(rw, T_Description).ClearContents
    Cells(rw, T_Description).ClearComments
    Cells(rw, T_Description).Value = DescriptionString
    Else 'as comment
    InsertComment DescriptionString, T_Description, True, rw
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
Sub SetDataValidation(Col As Integer, ValidationOptionStr As String, Optional rw As Integer)

If rw = 0 Then rw = Selection.Row

    With Cells(rw, Col).Validation
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
Dim rw As Integer

    'set row
    If InputRw = 0 Or IsMissing(InputRw) Then
    rw = Selection.Row
    Else
    rw = InputRw
    End If
    
'add comment with more detail
Set CheckRng = Cells(rw, Col)

    If Not CheckRng.Comment Is Nothing Then
        If append = False Then 'delete!
        CheckRng.Comment.Delete
        Else 'append, then clear what's in there
        CommentStr = CheckRng.Comment.text & " // " & CommentStr
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

