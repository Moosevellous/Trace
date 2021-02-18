Attribute VB_Name = "NewSheets"
'==============================================================================
'PUBLIC VARIABLES
'==============================================================================
Public ImportSheetName As String
Public ImportAsTabs As Boolean
Public Description() As String

'==============================================================================
' Name:     SheetExists
' Author:   PS
' Desc:     Returns TRUE or FALSE depending on if a worksheet exists
' Args:     WS_Name, a string that is the name of the worksheet
' Comments: (1) Appropriated from the internet
'==============================================================================
Private Function SheetExists(WS_Name As String) As Boolean
    Dim WS As Worksheet
    On Error Resume Next
    Set WS = Worksheets(WS_Name)
    If Not WS Is Nothing Then SheetExists = True
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BELOW HERE BE SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'==============================================================================
' Name:     New_Tab
' Author:   PS
' Desc:     Creates a new worksheet (tab) in the current workbook
' Args:     ImportName, the type of blank sheet to be brought in:
'           e.g. OCT TOA MECH etc.
' Comments: (1) Appropriated from the internet
'==============================================================================
Sub New_Tab(ImportName As String)

Dim TemplatePath As String
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

'fix for when the worksheet doesn't have focus and crashes.
    If Cells(1, 1).Value = "" Then
    Cells(1, 1).Value = "x"
    Cells(1, 1).ClearContents
    End If
DoEvents

NumSheets = ActiveWorkbook.Sheets.Count
TemplatePath = TEMPLATELOCATION & "\" & ImportName & ".xlsm"

'suppress error messages
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

'Test the file location
'Debug.Print TemplatePath
    If Len(Dir(TemplatePath)) = 0 Then
    msg = MsgBox("Error - File not found!", vbOKOnly, "LOST IN SPACE?!")
    Exit Sub
    End If
    
'open file
Workbooks.Open fileName:=TemplatePath, ReadOnly:=True, Notify:=False
Application.DisplayAlerts = True
Application.AskToUpdateLinks = True
DoEvents

Set TemplateBook = ActiveWorkbook
    'loop through all sheets and copy
    For sh = 1 To TemplateBook.Sheets.Count
    Set TemplateSheet = TemplateBook.Sheets(sh)
    TemplateSheet.Copy After:=CurrentBook.Sheets(NumSheets)
    DoEvents
    Next sh
    
'resume normal error messages
Application.DisplayAlerts = False
Workbooks(CurrentBook.Name).Styles.Merge (TemplateBook.Name) 'merge styles
Application.DisplayAlerts = True

TemplateBook.Close SaveChanges:=False
DoEvents

    'there are buttons on this sheet that need to be repointed
    If ImportName = "CVT" Then ReassignConversionButtons

Exit Sub

'catch errrors
errors:
Debug.Print Err.Number; " - "; Err.Description

    Select Case Err.Number
    Case Is = 1004
    msg = MsgBox("Not enough rows in workbook: " & chr(10) & CurrentBook.Name & _
        chr(10) & "Convert workbook to XLSX format and try again.", _
        vbOKOnly, "XLS error")
    TemplateBook.Close SaveChanges:=False
    End Select

End Sub

'==============================================================================
' Name:     ReassignConversionButtons
' Author:   PS
' Desc:     Repoints buttons to macros in current Trace location
' Args:     None
' Comments: (1) Called when inserting a new sheet
'==============================================================================
Sub ReassignConversionButtons()
Dim C2O_Action As String
Dim C2O_Action_new As String
Dim SplitStr() As String
    With ActiveSheet
    C2O_Action = .Shapes("btnConvertToOctaves").OnAction
    'Debug.Print C2O_Action
    SplitStr = Split(C2O_Action, "!", -1, vbTextCompare)
    C2O_Action_new = "'" & Application.AddIns("Trace").FullName & "'!" & SplitStr(1)
    'Debug.Print C2O_Action_new
    .Shapes("btnConvertToOctaves").OnAction = C2O_Action_new
    End With
End Sub

'==============================================================================
' Name:     SameType
' Author:   PS
' Desc:     Opens another sheet of the same type as the current sheet
' Args:     None
' Comments: (1)
'==============================================================================
Sub Same_Type()
Dim TemplateBook As Workbook
Dim CurrentBook As Workbook
Dim TemplateSheet As Worksheet
Dim NumSheets As Long
Dim i As Integer
Dim TypeCode As String
Dim OpenPathStr As String

On Error GoTo errors

Application.ScreenUpdating = False

GetSettings 'get all variable information

    If IsNamedRange("TYPECODE") = True Then
    TypeCode = Range("TYPECODE").Value
    Else
    msg = MsgBox("No sheet type selected, perhaps try adding a new one?", _
        vbOKOnly, "Oh sheet...")
    Exit Sub
    End If
    
Set CurrentBook = ActiveWorkbook
NumSheets = ActiveWorkbook.Sheets.Count
OpenPathStr = TEMPLATELOCATION & "\" & TypeCode

'suppress error messages
Application.DisplayAlerts = False
Application.EnableEvents = False
Application.AskToUpdateLinks = False

'open
Workbooks.Open (OpenPathStr)  'public variable
DoEvents
Set TemplateBook = ActiveWorkbook
Application.DisplayAlerts = True
Application.AskToUpdateLinks = True

'copy in
TemplateBook.Sheets(1).Copy After:=CurrentBook.Sheets(NumSheets)
DoEvents

'close
TemplateBook.Close SaveChanges:=False
    
    'there are buttons on this sheet that need to be repointed
    If TypeCode = "CVT" Then ReassignConversionButtons

'resume normal error messages
Application.DisplayAlerts = True
Application.EnableEvents = True

'catch errrors
errors:
Debug.Print Err.Number; " - "; Err.Description

    If Left(Err.Description, 5) = "Sorry" Then 'file not found
    ErrorOCTTOOnly
    End If

End Sub

'==============================================================================
' Name:     LoadCalcFieldSheet
' Author:   PS
' Desc:     Loads Field Sheets, Equipment Import Sheets, and Standard
'           Calculation sheets
' Args:     ImportSheetType, string which is either Standard, Field or
'           EquipmentImport, set by the ribbon callback function
' Comments: (1)
'==============================================================================
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
'for user form
Dim CheckBoxSpacer As Integer
Dim DefaultCurrentTop As Integer
Dim DefaultCheckboxWidth As Integer
Dim CurrentTop As Integer
Dim CurrentCol As Integer
Dim BottomBuffer As Integer
Dim optControl As control

GetSettings 'get all variable information

'set the object Current book. If it doesn't exist, create one!
Set CurrentBook = ActiveWorkbook
    If CurrentBook Is Nothing Then
    Application.Workbooks.Add
    DoEvents
    Set CurrentBook = ActiveWorkbook
    End If
CurrentBookName = CurrentBook.Name

Application.StatusBar = "Generating list of templates..."
Set fso = CreateObject("Scripting.FileSystemObject")

    'select type of import
    If ImportSheetType = "Standard" Then
        Set ScanFolder = fso.GetFolder(STANDARDCALCLOCATION)
        frmStandardCalc.Caption = "Standard Calculation Sheets"
    ElseIf ImportSheetType = "Field" Then
        Set ScanFolder = fso.GetFolder(FIELDSHEETLOCATION)
        frmStandardCalc.Caption = "Field Sheets"
    ElseIf ImportSheetType = "EquipmentImport" Then
        Set ScanFolder = fso.GetFolder(EQUIPMENTSHEETLOCATION)
        frmStandardCalc.Caption = "Equipment Import Sheets"
    End If

'set layout variables
CheckBoxSpacer = 20 'px
DefaultCurrentTop = 10 'px: space at top of each column
CurrentTop = DefaultCurrentTop 'gets reset throughout
DefaultCheckboxWidth = frmStandardCalc.mPageSheets.Width / 2 'two columns
numbookmarks = 1
CurrentCol = 0
CurrentPage = 0
BottomBuffer = 50 'px: space at the bottom

    'create radio buttons
    For Each standardSheet In ScanFolder.Files
    'Debug.Print standardSheet.Name
        If Left(standardSheet.Name, 1) <> "~" Then
        'CheckColumn
        Set optControl = frmStandardCalc.mPageSheets.Pages(CurrentPage) _
                        .Controls.Add("Forms.OptionButton.1")
            With optControl
            .Caption = standardSheet.Name
            .Top = CurrentTop
            .Left = 5 + (CurrentCol * DefaultCheckboxWidth)
            .Width = DefaultCheckboxWidth
            End With
            
        CurrentTop = CurrentTop + CheckBoxSpacer
        
            'check for second column
            If CurrentTop > frmStandardCalc.mPageSheets.Height - BottomBuffer Then
            CurrentTop = DefaultCurrentTop
            CurrentCol = CurrentCol + 1
            End If
            'check for multipage
            If CurrentCol > 1 Then 'Note: starts at 0
            frmStandardCalc.mPageSheets.Pages.Add
            CurrentPage = CurrentPage + 1
            CurrentTop = DefaultCurrentTop
            CurrentCol = 0
            End If
        numbookmarks = numbookmarks + 1
        End If
    Next standardSheet

'Prompt the user
frmStandardCalc.Show
Application.StatusBar = False

    If btnOkPressed = False Then
    End
    End If

Application.StatusBar = "Opening " & ImportSheetName

If ImportSheetName = "" Then Exit Sub

'public variables set during setup, other variable returned from form
'Update Links set to False for a smoother experience
    If ImportSheetType = "Standard" Then
    Workbooks.Open STANDARDCALCLOCATION & "\" & ImportSheetName, False
    ElseIf ImportSheetType = "Field" Then
    Workbooks.Open FIELDSHEETLOCATION & "\" & ImportSheetName, False
    ElseIf ImportSheetType = "EquipmentImport" Then
    Workbooks.Open EQUIPMENTSHEETLOCATION & "\" & ImportSheetName, False
    End If

DoEvents
Application.StatusBar = False
Set TemplateBook = ActiveWorkbook

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'<----TODO Element types sheet for glazing, facade break-in etc
'<----BUT is it required if we implement databases?
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
'52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)
    Select Case TemplateBook.FileFormat
    Case Is = 51
    FilterIndexValue = 1
    Case Is = 52
    FilterIndexValue = 2
    End Select

    If ImportAsTabs = False Then
    Application.StatusBar = "Saving sheet..."
    SaveSheetAs_DateStamped TemplateBook.Name, FilterIndexValue
    Else
    Application.StatusBar = "Importing...."
    LastSheet = Workbooks(CurrentBookName).Sheets.Count
    Workbooks(ImportSheetName).Sheets.Select
    Sheets().Copy After:=Workbooks(CurrentBookName).Sheets(LastSheet)
    Workbooks(TemplateBook.Name).Close (False)
    End If

DoEvents

catch:
Application.StatusBar = False

End Sub

'==============================================================================
' Name:     SaveSheetAs_DateStamped
' Author:   PS
' Desc:     Saves as with reverse datestamp in title
' Args:     SaveAsName (string of default name), FilterIndex(For different
'           file types)
' Comments: (1)Called from LoadCalcFieldSheet
'==============================================================================
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
