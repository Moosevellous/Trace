Attribute VB_Name = "ImportFrom"
Sub IMPORT_FANTECH_DATA(SheetType As String)

On Error GoTo errHandler

Dim splitStr() As String
Dim RawBookName As String
Dim WriteRw As Integer
Dim PercentDone As Single

    If Left(SheetType, 2) = "TO" Then
    msg = MsgBox("Fantech import is not available for one-third octave band sheets", vbOKOnly, "Import: Impossible")
    End
    End If


Application.DefaultFilePath = ActiveWorkbook.Path
WorkbookName = ActiveWorkbook.Name
WriteRw = Selection.Row
WriteCol = 2
doneFiles = 0
runonce = False

Application.ScreenUpdating = False


'Open Raw File
File = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx),*.xlsx", _
ButtonText:="Please select file (in XLSX format)...", _
MultiSelect:=True)

numFiles = UBound(File)

    For fnumber = 1 To UBound(File)
    Workbooks.Open File(fnumber)
    DoEvents
    RawBookName = ActiveWorkbook.Name
    Application.StatusBar = RawBookName
    'Inlet
    FanType = Cells(7, 2).Value
    Range("B33:B40").Copy
    Workbooks(WorkbookName).Activate
    Cells(WriteRw, 2) = FanType & " - Inlet"
    Cells(WriteRw, 6).PasteSpecial Paste:=xlValues, Transpose:=True
    WriteRw = WriteRw + 1
    
    'Outlet
    Workbooks(RawBookName).Activate
    Range("B42:B49").Copy
    Workbooks(WorkbookName).Activate
    Cells(WriteRw, 2) = FanType & " - Outlet"
    Cells(WriteRw, 6).PasteSpecial Paste:=xlValues, Transpose:=True
        
    'Breakout
    'TODO BREAKOUT
    
    Application.CutCopyMode = False
    Workbooks(RawBookName).Close (False)
    
    'Status
    doneFiles = doneFiles + 1
    PercentDone = (doneFiles / numFiles)
    Application.ScreenUpdating = True
    'Cells(2, 11).Value = PercentDone
    Application.ScreenUpdating = False
    WriteRw = WriteRw + 1
    Next fnumber
    
Application.StatusBar = False
msg = MsgBox("Done!", vbOKOnly, "Macro Complete")

Exit Sub

errHandler:
  MsgBox "Error " & Err.Number & ": " & Err.DESCRIPTION, vbOKOnly, "Error"

End Sub

Sub Import_INSUL(SheetType As String)
Dim PasteCol As Integer
CheckRow (Selection.Row) 'CHECK FOR NON HEADER ROWS

    If Left(SheetType, 2) = "TO" Then
    Call ParameterUnmerge(Selection.Row, SheetType)
    
    ClipData = GetClipBoardText
    splitData = Split(ClipData, vbCr, Len(ClipData), vbTextCompare)
    PasteCol = 2
        For i = 0 To UBound(splitData)
        'split on tab
        splitline = Split(splitData(i), vbTab, Len(splitData(i)), vbTextCompare)
        Cells(Selection.Row, PasteCol).Value = splitline(UBound(splitline))
            If PasteCol = 2 Then
            TitleStr = splitline(UBound(splitline))
            PasteCol = 5 'skip straight to colume E
            ElseIf PasteCol = 25 Then
                If InStr(1, TitleStr, "FLOOR", vbTextCompare) = 0 Then 'not floor
                Cells(Selection.Row, 26).Value = "=RwRate(H" & Selection.Row & ":Y" & Selection.Row & ")" 'Rw
                Cells(Selection.Row, 26).NumberFormat = """Rw ""0"
                Cells(Selection.Row, 27).Value = "=CtrRate(H" & Selection.Row & ":Y" & Selection.Row & ",Z" & Selection.Row & ")" 'Ctr
                Cells(Selection.Row, 27).NumberFormat = ";Ct\r -0;"
                End If
            End
            Else
            PasteCol = PasteCol + 1
            End If
        Next i
    Else
    msg = MsgBox("Implemented for one-third octave sheet only. Please try again later.", vbOKOnly, "To be continued...")
    End If
End Sub

Function GetClipBoardText()
   Dim DataObj As MsForms.DataObject
   Set DataObj = New MsForms.DataObject

   On Error GoTo Whoa

   '~~> Get data from the clipboard.
   DataObj.GetFromClipboard

   '~~> Get clipboard contents
   GetClipBoardText = DataObj.GetText(1)

   Exit Function
Whoa:
   If Err <> 0 Then MsgBox "Data on clipboard is not text or is empty"
End Function

