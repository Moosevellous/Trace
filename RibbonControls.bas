Attribute VB_Name = "RibbonControls"
Function NamedRangeExists(strRangeName As String) As Boolean
Dim rngExists  As Range
On Error Resume Next
Set rngExists = Range(strRangeName)
NamedRangeExists = True
    If rngExists Is Nothing Then
    NamedRangeExists = False
    msg = MsgBox("Error: Named Range TYPECODE missing!" & chr(10) & chr(10) & "Description: Trace functions require that you use a blank calculation sheet. " & chr(10) & "Try clicking 'Add Sheet' in the Load group on the Trace Ribbon (top left).", vbOKOnly, "Sorryyyyyyyyyy")
    End If
    On Error GoTo 0
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'LOAD
Sub btnLoad(control As IRibbonControl)
    New_Tab
End Sub

Sub btnSameType(control As IRibbonControl)
    Same_Type
End Sub

Sub btnStandardCalc(control As IRibbonControl)
    LoadCalcFieldSheet ("Standard")
End Sub

Sub btnFieldSheet(control As IRibbonControl)
    LoadCalcFieldSheet ("Field")
End Sub

Sub btnEquipmentImport(control As IRibbonControl)
    LoadCalcFieldSheet ("EquipmentImport")
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'THE BASICS (NOT INCLUDING WALLY DE BACKER)
Sub btnSPLSUM(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then
    InsertBasicFunction ActiveSheet.Range("TYPECODE").Value, "SPLSUM"
    End If
End Sub

Sub btnSPLMINUS(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then
    InsertBasicFunction ActiveSheet.Range("TYPECODE").Value, "SPLMINUS"
    End If
End Sub

Sub btnSPLAV(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then
    InsertBasicFunction ActiveSheet.Range("TYPECODE").Value, "SPLAV"
    End If
End Sub

Sub btnSPLSUMIF(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then
    InsertBasicFunction ActiveSheet.Range("TYPECODE").Value, "SPLSUMIF"
    End If
End Sub

Sub btnSPLAVIF(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then
    InsertBasicFunction ActiveSheet.Range("TYPECODE").Value, "SPLAVIF"
    End If
End Sub

Sub btnWavelength(control As IRibbonControl)
DoesNotExist
End Sub

Sub btnSpeedOfSound(control As IRibbonControl)
DoesNotExist
End Sub

Sub btnFrequencyBandCutoff(control As IRibbonControl)
DoesNotExist
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IMPORT FROM
Sub btnImportFantech(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then IMPORT_FANTECH_DATA (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnImportInsul(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then Import_INSUL (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnImportZorba(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then Import_Zorba (ActiveSheet.Range("TYPECODE").Value)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ROW FUNCTIONS

Sub btnClearRw(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then ClearRw (ActiveSheet.Range("TYPECODE").Value)
End Sub

'Sub btnAWeight(control As IRibbonControl)
'On Error Resume Next
'A_weight_oct (ActiveSheet.Range("TYPECODE").Value)
'End Sub

Sub btnFlipSign(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then FlipSign (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnMoveUp(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then MoveUp (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnMoveDown(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then MoveDown (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnSingleCorrection(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then SingleCorrection (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnAutoSum(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then AutoSum (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnManual_ExtendFunction(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then Manual_ExtendFunction (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnTenLogN(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then TenLogN (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnTenLogOneOnT(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then TenLogOneOnT (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnOneThirdsToOctave(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then OneThirdsToOctave (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnConvertAWeight(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then ConvertToAWeight (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnConvertCWeight(control As IRibbonControl)
    DoesNotExist
End Sub

Sub btnRowReference(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then RowReference (ActiveSheet.Range("TYPECODE").Value)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NOISE FUNCTIONS

'Sub btnAirAbsorption(control As IRibbonControl)
'    If NamedRangeExists("TYPECODE") Then AirAbsorption (ActiveSheet.Range("TYPECODE").Value)
'End Sub

Sub btnISO_full(control As IRibbonControl)
If NamedRangeExists("TYPECODE") Then ISO_full (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnISO_Adiv(control As IRibbonControl)
If NamedRangeExists("TYPECODE") Then A_div (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnISO_Aatm(control As IRibbonControl)
If NamedRangeExists("TYPECODE") Then A_atm (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnISO_Agr(control As IRibbonControl)
DoesNotExist
End Sub

Sub btnISO_Abar(control As IRibbonControl)
DoesNotExist
End Sub

Sub btnISO_Cmet(control As IRibbonControl)
DoesNotExist
End Sub


Sub btnDistance(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then Distance (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnDistanceLine(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then DistanceLine (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnDistancePlane(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then DistancePlane (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnAreaCorrection(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then Area (ActiveSheet.Range("TYPECODE").Value)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub btnRegenNoise(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then RegenNoise (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnDirectReverberantSum(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then DirRevSum (ActiveSheet.Range("TYPECODE").Value)
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''
    'MECH SUBGROUP
    Sub btnDuctAtten(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then DuctAtten (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFlexDuct(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then FlexDuct (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnERL(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then ERL (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnElbow(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then ElbowLoss (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnDuctSplit(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then DuctSplit (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnSilencer(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then Silencer (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnLouvres(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then Louvres (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnPlenum(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then Plenum (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnDuctBreakout(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then DuctBreakout (ActiveSheet.Range("TYPECODE").Value)
    End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''
    'ROOM LOSS SUBGROUP
    
    Sub btnRoomLoss(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then RoomLoss (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnRoomLossRC(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then RoomLossRC (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnRoomLossRT(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then RoomLossRT (ActiveSheet.Range("TYPECODE").Value)
    End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Sub btnceilingIL(control As IRibbonControl)
        If NamedRangeExists("TYPECODE") Then DoesNotExist
    End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CURVE FUNCTIONS

Sub btnNRcurve(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutNR (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnNCcurve(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutNC (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnPNCcurve(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutPNC (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnRwCurve(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutRw (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnSTCCurve(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutSTC (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnLnwCurve(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutLnw (ActiveSheet.Range("TYPECODE").Value)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ESTIMATOR FUNCTIONS

Sub btnFanSimple(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutLwFanSimple (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnPumpSimple(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutLwPumpSimple (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnCoolingTower(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutLwCoolingTower (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnCompressor(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutCompressorSmall (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnElectricMotor(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutElectricMotorSmall (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnGasTurbine(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutGasTurbine (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnSteamTurbine(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutSteamTurbine (ActiveSheet.Range("TYPECODE").Value)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VIBRATION

Sub btnMMPS2DB(control As IRibbonControl) 'change this name
    If NamedRangeExists("TYPECODE") Then VibLin2DB (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnDB2MMPS(control As IRibbonControl) 'change this name
    If NamedRangeExists("TYPECODE") Then VibDB2Lin (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnVibConvert(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then VibConvert (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnASHRAEcurve(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then PutVCcurve (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnCouplingLoss(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then CouplingLoss (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnAmplification(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then BuildingAmplification (ActiveSheet.Range("TYPECODE").Value)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SHEET FUNCTIONS

Sub btnHeaderBlock(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then HeaderBlock (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnClearHeaderBlock(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then ClearHeaderBlock (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnFormatBorders(control As IRibbonControl)
FormatBorders
End Sub

Sub btnPlot(control As IRibbonControl)
    'If chart is selected then edit, otherwise start a new chart
    If ActiveChart Is Nothing Then
        If NamedRangeExists("TYPECODE") Then
        Plot (ActiveSheet.Range("TYPECODE").Value)
        End If
    Else
    frmChartFormatter.Show
    End If
End Sub

Sub btnHeatMap(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then HeatMap (ActiveSheet.Range("TYPECODE").Value)
End Sub

Sub btnFixReferences(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then FixReferences (ActiveSheet.Range("TYPECODE").Value)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FORMAT / STYLE
    
    ''''''
    'Styles
    ''''''
    Sub btnFmtTitle(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then fmtTitle (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFmtUnmiti(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then fmtUnmiti (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFmtMiti(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then fmtMiti (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFmtSource(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then fmtSource (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFmtSilencer(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then fmtSilencer (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFmtUserInput(SheetType As String)
    If NamedRangeExists("TYPECODE") Then fmtUserInput (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFmtComment(SheetType As String)
    If NamedRangeExists("TYPECODE") Then fmtComment (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFmtSubtotal(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then fmtSubtotal (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFmtTotal(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then fmtTotal (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnFmtReference(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then fmtReference (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    Sub btnfmtNormal(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then fmtNormal (ActiveSheet.Range("TYPECODE").Value)
    End Sub
    
    ''''''
    'Units
    ''''''
    Sub btnFmtUnitMetres(control As IRibbonControl)
    Unit_m Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub
    
    Sub btnFmtUnitMetresSquared(control As IRibbonControl)
    Unit_m2 Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub
    
    Sub btnFmtUnitMetresSquaredPerSecond(control As IRibbonControl)
    Unit_m2ps Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub
    
    Sub btnFmtUnitMetresCubedPerSecond(control As IRibbonControl)
    Unit_m3ps Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub
    
    Sub btnFmtUnitdB(control As IRibbonControl)
    Unit_dB Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub
    
    Sub btnFmtUnitdBA(control As IRibbonControl)
    Unit_dBA Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub
    
    Sub btnFmtUnitkW(control As IRibbonControl)
    Unit_kW Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub
    
    Sub btnFmtUnitPa(control As IRibbonControl)
    Unit_Pa Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub
    
    Sub btnFmtUnitQ(control As IRibbonControl)
    Unit_Q Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub
    
    Sub btnFmtClear(control As IRibbonControl)
    Unit_Clear Selection.Column, Selection.Column + Selection.Columns.Count - 1
    End Sub

Sub btnTarget(control As IRibbonControl)
    If NamedRangeExists("TYPECODE") Then Target (ActiveSheet.Range("TYPECODE").Value)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'HELP
Sub btnOnlineHelp(control As IRibbonControl)
GetHelp
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ERROR CATCHING
'named range for sheet type would enable error catching
'if named range throws error, can catch at button press
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ErrorTypeCode()
msg = MsgBox("Function only possible in template sheets.", vbOKOnly, "Waggling finger of shame")
End
End Sub

Sub DoesNotExist()
msg = MsgBox("Feature does not exist yet - please try again later", vbOKOnly, "Maybe one day....?")
End Sub

Sub ErrorOctOnly()
msg = MsgBox("Function only possible in octave bands.", vbOKOnly, "Once.....twice.....three times an octave")
End
End Sub

Sub ErrorThirdOctOnly()
msg = MsgBox("Function only possible in one-third octave bands.", vbOKOnly, "Fool me three times.....")
End
End Sub

Sub ErrorLFTOOnly()
msg = MsgBox("Function only possible in low-frequency one-third octave bands.", vbOKOnly, "All about that bass")
End
End Sub
