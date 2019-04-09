Attribute VB_Name = "VARS"
''''''''''''''''''''''
'GLOBAL VARIABLES'''''
''''''''''''''''''''''
Public TEMPLATELOCATION As String
Public STANDARDCALCLOCATION As String
Public FIELDSHEETLOCATION As String
Public EQUIPMENTSHEETLOCATION As String
Public ASHRAE_DUCT_TXT As String
Public ASHRAE_FLEX As String
Public ASHRAE_REGEN As String
Public ENGINEER As String
Public PROJECTNO As String
Public PROJECTNAME As String
Public PROJECTINFODIRECTORY As String
Public FANTECH_SILENCERS As String '<---TODO remove reference to Fantech?
Public ACOUSTIC_LOUVRES As String
Public colourUSERINPUT As Long
Public colourFINALRESULT  As Long


''''''''''''''''''''''
'END GLOBAL VARIABLES
''''''''''''''''''''''

'sub to initialise
Public Sub GetSettings()
Dim RootPath  As String

On Error Resume Next

'Debug.Print "No of addins: " & Application.AddIns.Count
'    For Each ad In Application.AddIns
'    Debug.Print ad.Name
'        If ad.Name = "NoiseCalc.xlam" Then
'        RootPath = Application.AddIns("NoiseCalc").Path
'        End If
'    Next ad

    If Application.AddIns.count = 0 Then 'catches the error
    RootPath = "U:\SectionData\Property\Specialist Services\Acoustics\1 - Technical Library\Excel Add-in\Trace" 'hard coded location of AddIn as a fallback
    Else
    RootPath = Application.AddIns("Trace").Path
    End If

'Debug.Print RootPath

TEMPLATELOCATION = RootPath & "\Template Sheets\Blank Calculation Sheet.xlsm"
STANDARDCALCLOCATION = RootPath & "\Standard Calc Sheets"
FIELDSHEETLOCATION = RootPath & "\Field Sheets"
EQUIPMENTSHEETLOCATION = RootPath & "\Equipment Import Sheets"
ASHRAE_DUCT_TXT = RootPath & "\ASHRAE DATA\ASHRAE_DUCTS.txt"
ASHRAE_FLEX = RootPath & "\ASHRAE DATA\ASHRAE_FLEX.txt"
ASHRAE_REGEN = RootPath & "\ASHRAE DATA\ASHRAE_REGEN.txt"
FANTECH_SILENCERS = RootPath & "\Silencers.txt"
ACOUSTIC_LOUVRES = RootPath & "\Louvres.txt"

'Colours
colourUSERINPUT = RGB(254, 253, 195) '<- TODO make into styles
colourFINALRESULT = RGB(146, 205, 220)
End Sub






