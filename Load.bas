Attribute VB_Name = "Load"
Public ImportSheetName As String
Public DESCRIPTION() As String

Private Function SheetExists(sWSName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sWSName)
    If Not ws Is Nothing Then SheetExists = True
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub New_Tab()

Dim TemplateBook As Workbook
Dim CurrentBook As Workbook
Dim TemplateSheet As Worksheet
Dim NumSheets As Long
Dim i As Integer

On Error GoTo errors:

GetSettings 'get all variable information

Set CurrentBook = ActiveWorkbook
    
If CurrentBook Is Nothing Then
    Application.Workbooks.Add
    DoEvents
    Set CurrentBook = ActiveWorkbook
End If

'dodgy fix?
    If Cells(1, 1).Value = "" Then
    Cells(1, 1).Value = "x"
    Cells(1, 1).ClearContents
    End If
DoEvents

NumSheets = ActiveWorkbook.Sheets.Count

Workbooks.Open Filename:=TEMPLATELOCATION, ReadOnly:=True 'global variable
DoEvents

Set TemplateBook = ActiveWorkbook
    With TemplateBook
    
    ReDim Preserve DESCRIPTION(.Sheets.Count)
    
        If frmLoadTemplate.cBoxSelectTemplate.ListCount = 0 Then 'only if list is not populated
            For i = 1 To .Sheets.Count
            frmLoadTemplate.cBoxSelectTemplate.AddItem .Sheets(i).Name
            DoEvents
            DESCRIPTION(i) = .Sheets(i).Cells(3, 15).Comment.Text 'multiline is on, paragraph marks work!
            Next i
        End If
        
            frmLoadTemplate.Show
            
        If SheetExists(ImportSheetName) Then
            Set TemplateSheet = .Sheets(ImportSheetName)
            TemplateSheet.Copy after:=CurrentBook.Sheets(NumSheets)
            DoEvents
        ElseIf ImportSheetName = "" Then 'Cancel Clicked
        TemplateBook.Close SaveChanges:=False
        DoEvents
        End
        Else
            MsgBox "There is no sheet with name " & ImportSheetName & "in:" & vbCr & .Name
        End If
    DoEvents
    End With

Application.DisplayAlerts = False 'suppress error message
ActiveWorkbook.Styles.Merge (TemplateBook.Name) 'merge styles
Application.DisplayAlerts = True 'but not the others


TemplateBook.Close SaveChanges:=False

Exit Sub

errors:

    Select Case Err.Number
    Case Is = 1004
    msg = MsgBox("Not enough rows in workbook: " & chr(10) & CurrentBook.Name & chr(10) & "Convert workbook to XLSX format and try again.", vbOKOnly, "XLS error")
    TemplateBook.Close SaveChanges:=False
    End Select

End Sub




Sub Same_Type()
Dim TemplateBook As Workbook
Dim CurrentBook As Workbook
Dim TemplateSheet As Worksheet
Dim NumSheets As Long
Dim i As Integer
Dim TypeCode As String

Application.ScreenUpdating = False

GetSettings 'get all variable information

    If NamedRangeExists("TYPECODE") = True Then
    TypeCode = Range("TYPECODE").Value
    Else
    msg = MsgBox("No sheet type selected, perhaps try adding a new one?", vbOKOnly, "Oh sheet...")
    Exit Sub
    End If

Set CurrentBook = ActiveWorkbook
    If CurrentBook Is Nothing Then
    Application.Workbooks.Add
    DoEvents
    Set CurrentBook = ActiveWorkbook
    End If
NumSheets = ActiveWorkbook.Sheets.Count

Workbooks.Open (TEMPLATELOCATION) 'global variable
DoEvents
Set TemplateBook = ActiveWorkbook
With TemplateBook

ReDim Preserve DESCRIPTION(.Sheets.Count)

    If SheetExists(TypeCode) Then
        Set TemplateSheet = .Sheets(TypeCode)
        TemplateSheet.Copy after:=CurrentBook.Sheets(NumSheets)
        DoEvents
    ElseIf ImportSheetName = "" Then 'Cancel Clicked
    TemplateBook.Close SaveChanges:=False
    DoEvents
    End
    Else
        MsgBox "There is no sheet with name " & TypeCode & "in:" & vbCr & .Name
    End If
DoEvents
End With

TemplateBook.Close SaveChanges:=False

End Sub


Sub LoadCalcFieldSheet(ImportSheetType As String)

'On Error GoTo catch:

Dim TemplateBook As Workbook
Dim CurrentBook As Workbook
Dim CurrentBookName As String
Dim TemplateSheet As Worksheet
Dim NumSheets As Long
Dim i As Integer
Dim fso As FileSystemObject
Dim ScanFolder As Folder
Dim standardSheet As File
Dim FileNameStr As String
Dim FilterIndexValue As Integer

GetSettings 'get all variable information

Set CurrentBook = ActiveWorkbook
    If CurrentBook Is Nothing Then
    Application.Workbooks.Add
    DoEvents
    Set CurrentBook = ActiveWorkbook
    CurrentBookName = CurrentBook.Name
    End If

Application.StatusBar = "Generating list of templates..."
Set fso = CreateObject("Scripting.FileSystemObject")

    'select type of import
    If ImportSheetType = "Standard" Then
    Set ScanFolder = fso.GetFolder(STANDARDCALCLOCATION)
    ElseIf ImportSheetType = "Field" Then
    Set ScanFolder = fso.GetFolder(FIELDSHEETLOCATION)
    ElseIf ImportSheetType = "EquipmentImport" Then
    Set ScanFolder = fso.GetFolder(EQUIPMENTSHEETLOCATION)
    End If


    If frmStandardCalc.cBoxSelectTemplate.ListCount <= 0 Then
        For Each standardSheet In ScanFolder.Files
            If Left(standardSheet.Name, 1) <> "#" Then
            frmStandardCalc.cBoxSelectTemplate.AddItem (standardSheet.Name)
            Application.StatusBar = "Loading: " & standardSheet.Name
            End If
        Next standardSheet
    End If

Application.StatusBar = False
frmStandardCalc.Show

'<-------Form returns here

Application.StatusBar = False

    If btnOkPressed = False Then
    End
    End If

Application.StatusBar = "Opening " & ImportSheetName

If ImportSheetName = "" Then Exit Sub

    If ImportSheetType = "Standard" Then
    Workbooks.Open (STANDARDCALCLOCATION & "\" & ImportSheetName) 'Global variable set from form code
    ElseIf ImportSheetType = "Field" Then
    Workbooks.Open (FIELDSHEETLOCATION & "\" & ImportSheetName) 'Global variable set from form code
    ElseIf ImportSheetType = "EquipmentImport" Then
    Workbooks.Open (EQUIPMENTSHEETLOCATION & "\" & ImportSheetName) 'Global variable set from form code
    End If

DoEvents
Application.StatusBar = False
Set TemplateBook = ActiveWorkbook

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TODO Element types insertion for glazing, facade break-in etc
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Select Case TemplateBook.FileFormat
    Case Is = 51
    FilterIndexValue = 1
    Case Is = 52
    FilterIndexValue = 2
    End Select
'51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
'52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)

    If ImportSheetType = "Standard" Or ImportSheetType = "EquipmentImport" Then 'Only save the Equipment Import sheets
    newTabs = MsgBox("Do you want to add to existing workbook '" & CurrentBookName & "'?" & _
    chr(10) & "Note: Clicking 'No' will Save As", vbYesNo, "Import as new tabs?")
        If newTabs = vbNo Then
        Application.StatusBar = "Saving sheet..."
        SaveSheetAs_DateStamped TemplateBook.Name, FilterIndexValue
        Else
        Application.StatusBar = "Importing...."
        ImportSheetAsNewTabs TemplateBook.Name, CurrentBook.Name
        Workbooks(TemplateBook.Name).Close (False)
        End If
    ElseIf ImportSheetType = "EquipmentImport" Then
    
    End If

DoEvents

catch:
Application.StatusBar = False

End Sub

Sub SaveSheetAs_DateStamped(SaveAsName As String, FilterIndex As Integer)

FileNameStr = Format(CStr(Now), "yyyymmdd") & " " & SaveAsName

    Ret = Application.GetSaveAsFilename(InitialFileName:=FileNameStr, _
            FileFilter:="Excel Macro Free Workbook (*.xlsx), *.xlsx," & _
            "Excel Macro Enabled Workbook (*.xlsm), *.xlsm,", _
            FilterIndex:=FilterIndex, _
            Title:="Save As")

    If Ret <> False Then
    ActiveWorkbook.SaveAs (Ret)
    End If
End Sub

Sub ImportSheetAsNewTabs(ImportBookName As String, MainBookName As String)

'Dim SheetsNameArray() As String
Dim LastSheet As Integer
LastSheet = Workbooks(MainBookName).Sheets.Count
Workbooks(ImportBookName).Sheets.Select
Sheets().Copy after:=Workbooks(MainBookName).Sheets(LastSheet)
End Sub

Sub Error_Catcher(e)

    Select Case e
    Case Is = 1004
    msg = MsgBox("Not enough rows", vbOKOnly, "XLS error")
    End Select
    
End Sub
