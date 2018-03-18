Attribute VB_Name = "FormatStyles"
Sub FormatAs_CellReference(SheetType As String)
Call GetSettings
    If Left(SheetType, 3) = "OCT" Then
    Range(Cells(Selection.Row, 2), Cells(Selection.Row, 15)).Font.Color = fmtREFERENCE
    ElseIf Left(SheetType, 2) = "TO" Then
    Range(Cells(Selection.Row, 2), Cells(Selection.Row, 25)).Font.Color = fmtREFERENCE
    End If
End Sub

Sub FormatAs_Total()
Call GetSettings
End Sub

Sub FormatAs_UserInput()
Call GetSettings
End Sub
