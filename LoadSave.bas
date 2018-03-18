Attribute VB_Name = "LoadSave"
Public IMPORTSHEETNAME As String
Public DESCRIPTION() As String

Sub New_Tab()

Dim TemplateBook As Workbook
Dim CurrentBook As Workbook
Dim TemplateSheet As Worksheet
Dim NumSheets As Long
Dim i As Integer

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
            
        If SheetExists(IMPORTSHEETNAME) Then
            Set TemplateSheet = .Sheets(IMPORTSHEETNAME)
            TemplateSheet.Copy After:=CurrentBook.Sheets(NumSheets)
            DoEvents
        ElseIf IMPORTSHEETNAME = "" Then 'Cancel Clicked
        TemplateBook.Close SaveChanges:=False
        DoEvents
        End
        Else
            MsgBox "There is no sheet with name " & IMPORTSHEETNAME & "in:" & vbCr & .Name
        End If
    DoEvents
    End With

TemplateBook.Close SaveChanges:=False

End Sub

Private Function SheetExists(sWSName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sWSName)
    If Not ws Is Nothing Then SheetExists = True
End Function


Sub Same_Type()
Dim TemplateBook As Workbook
Dim CurrentBook As Workbook
Dim TemplateSheet As Worksheet
Dim NumSheets As Long
Dim i As Integer
Dim TypeCode As String

Application.ScreenUpdating = False

GetSettings 'get all variable information
    If RangeExists("TYPECODE") Then
    TypeCode = Range("TYPECODE").Value
    Else
    End
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
        TemplateSheet.Copy After:=CurrentBook.Sheets(NumSheets)
        DoEvents
    ElseIf IMPORTSHEETNAME = "" Then 'Cancel Clicked
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

Function RangeExists(s As String) As Boolean
On Error GoTo nope
    RangeExists = Range(s).Count > 0
nope:
End Function


Sub StandardCalc()

'On Error GoTo catch:

Dim TemplateBook As Workbook
Dim CurrentBook As Workbook
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
    End If
NumSheets = ActiveWorkbook.Sheets.Count

Application.StatusBar = "Generating list of templates..."
Set fso = CreateObject("Scripting.FileSystemObject")
Debug.Print STANDARDCALCLOCATION
Set ScanFolder = fso.GetFolder(STANDARDCALCLOCATION)

    If frmStandardCalc.cBoxSelectTemplate.ListCount <= 0 Then
    For Each standardSheet In ScanFolder.Files
        If Left(standardSheet.Name, 1) <> "#" Then
        frmStandardCalc.cBoxSelectTemplate.AddItem (standardSheet.Name)
        End If
    Next standardSheet
    End If
    
frmStandardCalc.Show

Application.StatusBar = False

    If btnOkPressed = False Then
    End
    End If

Application.StatusBar = "Opening " & IMPORTSHEETNAME
Workbooks.Open (STANDARDCALCLOCATION & "\" & IMPORTSHEETNAME) 'Global variable set from form code
DoEvents
Application.StatusBar = False
Set TemplateBook = ActiveWorkbook

FileNameStr = Format(CStr(Now), "yyyymmdd") & " " & TemplateBook.Name
    Select Case TemplateBook.FileFormat
    Case Is = 51
    FilterIndexValue = 1
    Case Is = 52
    FilterIndexValue = 2
    End Select
'51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
'52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)

Application.StatusBar = "Saving sheet..."

Ret = Application.GetSaveAsFilename(InitialFileName:=FileNameStr, _
                                    FileFilter:="Excel Macro Free Workbook (*.xlsx), *.xlsx," & _
                                    "Excel Macro Enabled Workbook (*.xlsm), *.xlsm,", _
                                    FilterIndex:=FilterIndexValue, _
                                    Title:="Save As")

If Ret <> False Then
    ActiveWorkbook.SaveAs (Ret)
End If

DoEvents

catch:
Application.StatusBar = False

End Sub
