#AutoIt3Wrapper_Res_File_Add=compucom.jpg, rt_rcdata, TEST_JPG_1
#AutoIt3Wrapper_Add_Constants=n
#Region Includes
#include <resources.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>

#EndRegion Includes
#Region Variables
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
Global Const $xlWorkbookDefault = 51
$file = @ScriptDir & "\Report.XLS"
$sFilePath = ""
Global Const $xlExpression = 2
Global Const $xlAutomatic = -4105
Global Const $xlThemeColorDark1 = 1
Global Const $xlSolid = 1
Global Const $xlThemeColorAccent4 = 8
Global Const $xlDiagonalDown = 5
Global Const $xlNone = -4142
Global Const $xlDiagonalUp = 6
Global Const $xlEdgeLeft = 7
Global Const $xlContinuous = 1
Global Const $xlThin = 2
Global Const $xlEdgeTop = 8
Global Const $xlEdgeBottom = 9
Global Const $xlEdgeRight = 10
Global Const $xlInsideVertical = 11
Global Const $xlInsideHorizontal = 12
Global Const $xlEqual = 3
Global Const $xlCellValue = 1
Global Const $xlNotEqual = 4
Global Const $xlThemeColorAccent6 = 10
Global Const $xlTextString = 9
Global Const $xlContains = 0
Global Const $xlAnd = 1
Global Const $xlDown = -4121
Global Const $xlUp = -4162
Global Const $xlUnderlineStyleNone = -4142
Global Const $xlThemeColorLight1 = 2
Global Const $xlThemeFontMinor = 2


$filesarray = _FileListToArray(@ScriptDir, "*.txt")
$string = _ArrayToString($filesarray, "|", 1)
#EndRegion Variables
#Region GUI
GUICreate("Unix Audit App", 418, 100)
GUISetBkColor(0xFFFFFF)
$filename = GUICtrlCreateCombo("Choose File Name", 1, 1, 200, 200)
GUICtrlSetData(-1, $string, "")
$progress = GUICtrlCreateProgress(1, 27, 200, 20)
$labelout = GUICtrlCreateLabel("", 1, 50, 200, 20, 0x12)
GUICtrlSetColor(-1, 0xC0C0C0)
$processing = GUICtrlCreateLabel("", 3, 53, 194, 14)
GUICtrlSetColor(-1, 0xC0C0C0)
$button = GUICtrlCreateButton("Run", 1, 70, 100, 30, 0x4000)
$button2 = GUICtrlCreateButton("Close", 100, 70, 100, 30, 0x4000)
$pic1 = GUICtrlCreatePic("", 210, 22, 156, 88)
_ResourceSetImageToCtrl($pic1, "TEST_JPG_1")
GUISetState(@SW_SHOW)
GUISetState()
While 1
	Switch GUIGetMsg()
		Case $button
			$sFilePath = GUICtrlRead($filename)
			If $sFilePath <> "" And $sFilePath <> "Choose File Name" Then
				$sFilePath = @ScriptDir & "\" & $sFilePath
				Main()
				If StringInStr($file,".xlsx") = 0 Then
					FileDelete($file)
				EndIf
				GUICtrlSetData($progress, 100)
				GUICtrlSetData($processing, "Complete")
			Else
				MsgBox(0,"","Please choose a file to process.")
			EndIf
		Case $button2
			Exit
		Case $GUI_EVENT_CLOSE
			Exit
	EndSwitch
WEnd
#EndRegion GUI

Func Main()
	If FileExists($file) Then
		FileDelete($file)
	EndIf
	If FileExists($file & "x") Then
		FileDelete($file & "x")
	EndIf
	GUICtrlSetData($progress, 5)
	GUICtrlSetData($processing, "Cleaning Up Old Files")

	$oExcel = ObjCreate("Excel.Application")
	$oExcel.Application.DisplayAlerts = 0
	$oExcel.Visible = 0
	$oExcel.WorkBooks.Add
	$oExcel.ActiveSheet.Name = "Report"
	$oExcel.Range("A1").Select
	$oExcel.ActiveCell.FormulaR1C1 = "IP"
	$oExcel.Range("B1").Select
	$oExcel.ActiveCell.FormulaR1C1 = "Hosts"
	$oExcel.Range("C1").Select
	$oExcel.ActiveCell.FormulaR1C1 = "Master"
	$oExcel.Range("D1").Select
	$oExcel.ActiveCell.FormulaR1C1 = "AMPM"
	$oExcel.Range("E1").Select
	$oExcel.ActiveCell.FormulaR1C1 = "Clarify"
	$oExcel.Range("F1").Select
	$oExcel.ActiveCell.FormulaR1C1 = "ICMP"
	$oExcel.Range("G1").Select
	$oExcel.ActiveCell.FormulaR1C1 = "SNMP"
	$oExcel.Range("H1").Select
	$oExcel.ActiveCell.FormulaR1C1 = "Notes"
	$oExcel.Range("I1").Select
	$oExcel.ActiveCell.FormulaR1C1 = "OBJ Needed"

	$oExcel.ActiveWorkBook.SaveAs($file, -4143)
	$oExcel.Quit()
	GUICtrlSetData($progress, 10)
	GUICtrlSetData($processing, "Creating New Report")

	Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
	Local $sFileRead = FileRead($hFileOpen)
	$sFileRead = StringReplace($sFileRead, @TAB, "  ")
	$sFileRead = StringReplace($sFileRead, "not found", "not found")
	FileClose($hFileOpen)

	$split = StringSplit($sFileRead, '')

	For $x = 0 To UBound($split) - 1
		$split[$x] = $split[$x] & "::"
	Next
	$output = ParseToArray("::", $split, 2, 1, 1, "")
;~ 	_ArrayDisplay($output)

	GUICtrlSetData($processing, "Looping through lines")
	For $x = 1 To UBound($output) - 1
		$stepcalc = Int(($x / UBound($output))*100)
		$step = int(($stepcalc * .3)) + 10
		GUICtrlSetData($progress, $step)

		$result = StringRegExp($output[$x][0], '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', 3)
		If IsArray($result) Then
			$output[$x][0] = "Looking for " & $result[0]
		Else
			$output[$x][0] = "Error"
		EndIf
		$output[$x][1] = ""
		$output[$x][2] = Reg('(?<=\s{2})[a-zA-Z0-9-]*[^_][^\s](?=\s{2})', $output[$x][2])
		Do
			$output[$x][2] = StringReplace($output[$x][2], "  ", " ")
		Until StringInStr($output[$x][2], "  ") = 0
		$output[$x][3] = ""
		$output[$x][4] = Reg('SEEDNAME=[a-zA-Z0-9-]*[^_]', $output[$x][4])
		$output[$x][4] = StringStripWS($output[$x][4], 3)
		Do
			$output[$x][4] = StringReplace($output[$x][4], "  ", " ")
		Until StringInStr($output[$x][4], "  ") = 0
		$output[$x][5] = ""
		$output[$x][6] = Reg('DisplayName.*\s', $output[$x][6])
		$output[$x][7] = ""
		$output[$x][8] = Reg('(?<=Activated ).*(?=_)', $output[$x][8])
		$output[$x][9] = ""

		If StringInStr($output[$x][11], "ICMP") <> 0 Then
			$output[$x][13] = $output[$x][12]
			$output[$x][12] = $output[$x][11]
			If StringInStr($output[$x][10], "ICMPONLY") <> 0 Then
				$output[$x][11] = "ICMPONLY"
			Else
				$output[$x][11] = "Error"
			EndIf
		Else
			$output[$x][11] = StringReplace($output[$x][11], " = ", "~")
			$output[$x][11] = Reg('(?<=~).*', $output[$x][11])
		EndIf

		$output[$x][10] = ""
		$output[$x][12] = ""

		If StringInStr($output[$x][13], " 0%") <> 0 Then
			$output[$x][13] = "Pass"
		Else
			$output[$x][13] = "Error"
		EndIf
	Next

	$step += 1
	GUICtrlSetData($progress, $step)
	GUICtrlSetData($processing, "Creating Final Array")
#Region Final Array
	Dim $finalresults[UBound($output) + 1][7]
	$finalresults[0][0] = "IP"
	$finalresults[0][1] = "Hosts"
	$finalresults[0][2] = "Master"
	$finalresults[0][3] = "AMPM"
	$finalresults[0][4] = "Clarify"
	$finalresults[0][5] = "ICMP"
	$finalresults[0][6] = "SNMP"

	For $x = 1 To UBound($output) - 1
		$finalresults[$x][0] = StringReplace($output[$x][0], "Looking for ", "")
		$finalresults[$x][1] = $output[$x][2]
		$finalresults[$x][2] = StringReplace($output[$x][4], "SEEDNAME=", "")
		$finalresults[$x][3] = StringReplace($output[$x][6], "DisplayName = ", "")
		$finalresults[$x][4] = StringStripWS($output[$x][8],3)
		$finalresults[$x][5] = $output[$x][13]
		If StringInStr($output[$x][11], ".") <> 0 Then
			$finalresults[$x][6] = StringLeft($output[$x][11], StringInStr($output[$x][11], ".") - 1)
		Else
			$finalresults[$x][6] = $output[$x][11]
		EndIf

	Next

	$output = ""
#EndRegion Final Array
	$step += 5
	GUICtrlSetData($progress, $step)
	GUICtrlSetData($processing, "Writing Data to Excel")
#Region Populate Excel
	Global $conn = ObjCreate("ADODB.Connection")
	Global $DSN = ("Driver={Microsoft Excel Driver (*.xls)};DBQ=" & $file & ";readOnly=false;imex=1;HDR=Yes")
	$conn.Open($DSN)

	For $x = 1 To UBound($finalresults) - 1
		If $finalresults[$x][0] = "Error" And $finalresults[$x][1] = "Error" And $finalresults[$x][2] = "Error" And $finalresults[$x][3] = "Error" And $finalresults[$x][4] = "Error" And $finalresults[$x][5] = "Error" And $finalresults[$x][6] = "Error" Then
			ContinueLoop
		EndIf
		$query = 'insert into [Report$] ("IP", "Hosts", "Master", "AMPM", "Clarify", "ICMP", "SNMP") values (' & "'" & $finalresults[$x][0] & "','" & $finalresults[$x][1] & "','" & _
				$finalresults[$x][2] & "','" & $finalresults[$x][3] & "','" & $finalresults[$x][4] & "','" & $finalresults[$x][5] & "','" & $finalresults[$x][6] & "')"
;~ 	MsgBox(0,"",$query)
		$conn.execute($query)
	Next

	$conn.close
	$RS = ""
	$conn = ""
	$DSN = ""
#EndRegion Populate Excel
	$step += 5
	GUICtrlSetData($progress, $step)
	GUICtrlSetData($processing, "Formating Excel")
#Region Format Excel
	$oExcel = ObjCreate("Excel.Application")

	$oExcel.WorkBooks.Open($file)
	$oExcel.Application.DisplayAlerts = 0
	$oExcel.Visible = 0
	$iLastUsedCol = $oExcel.ActiveWorkbook.Sheets("Report").UsedRange.Columns.Count
	$iLastUsedRow = $oExcel.ActiveWorkbook.Sheets("Report").UsedRange.Rows.Count
	$range = $oExcel.Activesheet.Cells($iLastUsedRow, $iLastUsedCol).Address
	$range = StringReplace($range, "$", "")
	$range = "A1:" & $range
	#Region Color Alternating Rows
	$oExcel.Range($range).Select
	$oExcel.Selection.FormatConditions.Add($xlExpression, "", "=MOD(ROW(),2)=0", "")
	$oExcel.Selection.FormatConditions($oExcel.Selection.FormatConditions.Count).SetFirstPriority
	With $oExcel.Selection.FormatConditions(1).Interior
		.PatternColorIndex = $xlAutomatic
		.ThemeColor = $xlThemeColorDark1
		.TintAndShade = -0.14996795556505
	EndWith
	$oExcel.Selection.FormatConditions(1).StopIfTrue = True
	#EndRegion Color Alternating Rows
	#Region AutoFit Columns
	$oExcel.Range("A1:I1").Select
	$oExcel.Selection.Font.Bold = True
	$oExcel.Columns("A:A").EntireColumn.AutoFit
	$oExcel.Columns("B:B").EntireColumn.AutoFit
	$oExcel.Columns("C:C").EntireColumn.AutoFit
	$oExcel.Columns("D:D").EntireColumn.AutoFit
	$oExcel.Columns("E:E").EntireColumn.AutoFit
	$oExcel.Columns("F:F").EntireColumn.AutoFit
	$oExcel.Columns("G:G").EntireColumn.AutoFit
	$oExcel.Columns("H:H").EntireColumn.AutoFit
	$oExcel.Columns("I:I").EntireColumn.AutoFit
	#EndRegion AutoFit Columns

	#Region Create Borders
	$oExcel.Selection.Borders($xlDiagonalDown).LineStyle = $xlNone
	$oExcel.Selection.Borders($xlDiagonalUp).LineStyle = $xlNone
	With $oExcel.Selection.Borders($xlEdgeLeft)
		.LineStyle = $xlContinuous
		.ColorIndex = $xlAutomatic
		.TintAndShade = 0
		.Weight = $xlThin
	EndWith
	With $oExcel.Selection.Borders($xlEdgeTop)
		.LineStyle = $xlContinuous
		.ColorIndex = $xlAutomatic
		.TintAndShade = 0
		.Weight = $xlThin
	EndWith
	With $oExcel.Selection.Borders($xlEdgeBottom)
		.LineStyle = $xlContinuous
		.ColorIndex = $xlAutomatic
		.TintAndShade = 0
		.Weight = $xlThin
	EndWith
	With $oExcel.Selection.Borders($xlEdgeRight)
		.LineStyle = $xlContinuous
		.ColorIndex = $xlAutomatic
		.TintAndShade = 0
		.Weight = $xlThin
	EndWith
	With $oExcel.Selection.Borders($xlInsideVertical)
		.LineStyle = $xlContinuous
		.ColorIndex = $xlAutomatic
		.TintAndShade = 0
		.Weight = $xlThin
	EndWith
	With $oExcel.Selection.Borders($xlInsideHorizontal)
		.LineStyle = $xlContinuous
		.ColorIndex = $xlAutomatic
		.TintAndShade = 0
		.Weight = $xlThin
	EndWith
	#EndRegion Create Borders

	$oExcel.Range($range).Select

	$oExcel.Selection.FormatConditions.Add($xlCellValue, $xlEqual, "=""Pass""", "")
	$oExcel.Selection.FormatConditions($oExcel.Selection.FormatConditions.Count).SetFirstPriority
	With $oExcel.Selection.FormatConditions(1).Font
		.Color = -16752384
		.TintAndShade = 0
	EndWith
	With $oExcel.Selection.FormatConditions(1).Interior
		.PatternColorIndex = $xlAutomatic
		.Color = 13561798
		.TintAndShade = 0
	EndWith
	$oExcel.Selection.FormatConditions(1).StopIfTrue = False

	$oExcel.Selection.FormatConditions.Add($xlCellValue, $xlEqual, "=""Error""", "")
	$oExcel.Selection.FormatConditions($oExcel.Selection.FormatConditions.Count).SetFirstPriority
	With $oExcel.Selection.FormatConditions(1).Font
		.Color = -16383844
		.TintAndShade = 0
	EndWith
	With $oExcel.Selection.FormatConditions(1).Interior
		.PatternColorIndex = $xlAutomatic
		.Color = 13551615
		.TintAndShade = 0
	EndWith
	$oExcel.Selection.FormatConditions(1).StopIfTrue = False

	#Region Color Passes Green
	$oExcel.Range("G2:G" & $iLastUsedRow).Select
	$oExcel.Selection.FormatConditions.Add($xlCellValue, $xlNotEqual, "=""Error""", "")
	$oExcel.Selection.FormatConditions($oExcel.Selection.FormatConditions.Count).SetFirstPriority
	With $oExcel.Selection.FormatConditions(1).Font
		.Color = -16752384
		.TintAndShade = 0
	EndWith
	With $oExcel.Selection.FormatConditions(1).Interior
		.PatternColorIndex = $xlAutomatic
		.Color = 13561798
		.TintAndShade = 0
	EndWith
	$oExcel.Selection.FormatConditions(1).StopIfTrue = False
	#EndRegion Color Passes Green

	$oExcel.Range("I2").Select
	$oExcel.ActiveCell.FormulaR1C1 = "=COUNTIF(RC[-7]:RC[-2],""Error*"")"
	$oExcel.Range("I2").Select
	$oExcel.Selection.AutoFill($oExcel.Range("I2:I" & $iLastUsedRow))
	$oExcel.Range("I2:I" & $iLastUsedRow).Select
	$step += 5
	GUICtrlSetData($progress, $step)

	$step += 5
	GUICtrlSetData($progress, $step)
	GUICtrlSetData($processing, "Looking for errors to highlight")

	$oExcel.Range("$B$2:$B$" & $iLastUsedRow).Select
	$oExcel.Selection.AutoFilter
	$oExcel.ActiveSheet.Range("$B$2:$B$" & $iLastUsedRow).AutoFilter(1, "Error*")
	With $oExcel.Selection.Font
		.Color = -16383844
		.TintAndShade = 0
	EndWith
	With $oExcel.Selection.Interior
		.PatternColorIndex = $xlAutomatic
		.Color = 13551615
		.TintAndShade = 0
	EndWith
	$oExcel.Selection.AutoFilter

	$oExcel.Range("$C$2:$C$" & $iLastUsedRow).Select
	$oExcel.Selection.AutoFilter
	$oExcel.ActiveSheet.Range("$C$2:$C$" & $iLastUsedRow).AutoFilter(1, "Error*")
	With $oExcel.Selection.Font
		.Color = -16383844
		.TintAndShade = 0
	EndWith
	With $oExcel.Selection.Interior
		.PatternColorIndex = $xlAutomatic
		.Color = 13551615
		.TintAndShade = 0
	EndWith
	$oExcel.Selection.AutoFilter

	$range = $oExcel.Activesheet.Cells($iLastUsedRow, $iLastUsedCol).Address
	$range = StringReplace($range, "$", "")
	$range = "A1:" & $range
	$oExcel.Range($range).Select
	$oExcel.Selection.AutoFilter
	$oExcel.Range("A1").Select

	$oExcel.Range("A1:I1").Select
	With $oExcel.Selection.Interior
		.Pattern = $xlSolid
		.PatternColorIndex = $xlAutomatic
		.ThemeColor = $xlThemeColorAccent4
		.TintAndShade = 0.399975585192419
		.PatternTintAndShade = 0
	EndWith

	With $oExcel.Selection.Font
        .Name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = $xlUnderlineStyleNone
        .ThemeColor = $xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = $xlThemeFontMinor
    EndWith

	$oExcel.ActiveWorkBook.SaveAs($file & "X", $xlWorkbookDefault)
	$oExcel.Quit()
#EndRegion Format Excel

	$step += 5
	GUICtrlSetData($progress, $step)
	GUICtrlSetData($processing, "Wrapping Up")
EndFunc   ;==>Main

Func Reg($reg, $value)
	$result = StringRegExp($value, $reg, 3)
	If IsArray($result) Then
		If UBound($result) > 1 Then
			$value = "Error " & _ArrayToString($result, "   ")
		Else
			$value = $result[0]
		EndIf
	Else
		$value = "Error"
	EndIf
	Return $value
EndFunc   ;==>Reg

Func ParseToArray($string, $passedarray, $xeval, $yeval, $dateflag, $calling)
	$dimension = 1
	Do
		If StringInStr($passedarray[1], $string, 0, $dimension) <> 0 Then
			$dimension += 1
		ElseIf UBound($passedarray) < 3 Or UBound($passedarray, 2) > 1 Then

		ElseIf StringInStr($passedarray[2], $string, 0, $dimension) <> 0 Then
			$dimension += 1
		EndIf
	Until StringInStr($passedarray[1], $string, 0, $dimension) = 0
	$dimension -= 1
	If $dimension < 2 Then
		$dimension = 2
	EndIf
	$dimension = 14
;~ 	MsgBox(0, $x, $dimension)
	Dim $array[UBound($passedarray)][$dimension]
	For $xcounter = 0 To UBound($passedarray) - 1

		If $xcounter = 0 Then
			$array[$xcounter][0] = UBound($passedarray)
			$array[$xcounter][1] = $dimension
			ContinueLoop
		EndIf
;~ 		MsgBox(0,"",StringRight($passedarray[$xcounter], 1))
		$passedarray[$xcounter] = StringStripWS($passedarray[$xcounter], 3)
		If StringRight($passedarray[$xcounter], 1) <> $string Then
			$passedarray[$xcounter] = $passedarray[$xcounter] & $string
		EndIf
		If StringLeft($passedarray[$xcounter], 1) <> $string Then
			$passedarray[$xcounter] = $string & $passedarray[$xcounter]
		EndIf
		For $ycounter = 0 To $dimension - 1
			$start = ""
			$end = ""

			$start = StringInStr($passedarray[$xcounter], $string, 0, $ycounter + 1) + 1
			$end = StringInStr($passedarray[$xcounter], $string, 0, $ycounter + 2)
			If $start = 1 And $ycounter > 1 Then
				$end = 1
			EndIf
			If $end = 0 Then
				$end = StringLen($passedarray[$xcounter])
			EndIf
;~ 			MsgBox(0,$ycounter,"Start: " & $start & @CRLF & @CRLF & "End: " & $end)
			$difference = $end - $start
			$array[$xcounter][$ycounter] = StringMid($passedarray[$xcounter], $start, $difference)
			$array[$xcounter][$ycounter] = StringStripWS($array[$xcounter][$ycounter], 3)
			$dateflag = 0 ;Keep any date calculations from running.

		Next
	Next
;~ 	_ArrayDisplay($array)
	Return $array
EndFunc   ;==>ParseToArray

Func MyErrFunc()
	Local $HexNumber
	Local $strMsg

	$HexNumber = Hex($oMyError.Number, 8)
	$strMsg = "Error Number: " & $HexNumber & @CRLF
	$strMsg &= "WinDescription: " & $oMyError.WinDescription & @CRLF
	$strMsg &= "Script Line: " & $oMyError.ScriptLine & @CRLF
	MsgBox(0, "ERROR", $strMsg)
	SetError(1)
EndFunc   ;==>MyErrFunc

