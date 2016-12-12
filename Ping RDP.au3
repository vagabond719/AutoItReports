#include <Array.au3>
#include <Constants.au3>
#include <File.au3>

#region Constants
Global Const $xlWorkbookDefault = 51
Global Const $xlYes = 1
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
Global Const $msoThemeColorAccent1 = 5
Global Const $msoThemeColorText1 = 13
#endregion Constants

Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
$report = @ScriptDir & "\Report.xlsx"
$logFile = @ScriptDir & "\Error.log"

If FileExists($logFile) Then
	FileDelete($logFile)
	_FileCreate($logFile)
;~ 	Local $logFile = FileOpen(@ScriptDir & "\Error.log", 1)
Else
	_FileCreate($logFile)
;~ 	Local $logFile = FileOpen(@ScriptDir & "\Error.log", 1)
EndIf

If FileExists($report) Then
	FileDelete($report)
EndIf

#region Create Excel Report.xlsx
$oExcel = ObjCreate("Excel.Application")
$oExcel.Application.DisplayAlerts = 0
$oExcel.Visible = 0
$oExcel.WorkBooks.Add
$oExcel.ActiveSheet.Name = "Report"
$oExcel.Range("A1").Select
$oExcel.ActiveCell.FormulaR1C1 = "Server"
$oExcel.Range("B1").Select
$oExcel.ActiveCell.FormulaR1C1 = "RDP"
$oExcel.Range("C1").Select
$oExcel.ActiveCell.FormulaR1C1 = "Ping"
$oExcel.ActiveWorkBook.SaveAs($report, $xlWorkbookDefault)
$oExcel.Quit()
#endregion Create Excel Report.xlsx

ConsoleWrite("Start Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF)
ConsoleWrite(" ")
$conn = ObjCreate("ADODB.Connection")
$RS = ObjCreate("ADODB.Recordset")
$file = @ScriptDir & "\"
$DSN = ("Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & $file & ";")
$conn.Open($DSN)
$query = "select * from servers.txt"
$RS.open($query,$conn)
$strComputer = $rs.GetRows()
;~ _ArrayDisplay($strComputer)
$conn.close
$conn = ""
$RS = ""

Dim $final[UBound($strComputer)][3]

Global $conn = ObjCreate("ADODB.Connection")
$RS = ObjCreate("ADODB.Recordset")
Global $DSN = ("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DriverId=790;DBQ=" & $report & ";DefaultDir=" & @ScriptDir & ";IMEX=1;readonly=0;HDR=Yes")
$conn.Open($DSN)

For $x = 0 To UBound($strComputer) - 1
	$objWMIService = ObjGet("winmgmts:\\" & $strComputer[$x][0] & "\root\CIMV2")
	$started = ""
	If IsObj($objWMIService) Then
		$colItems = $objWMIService.ExecQuery("select Started from Win32_Service where name like 'Term%'")
		For $objItem In $colItems
		  $output = _GetDOSOutput("ping " & $strComputer[$x][0] & " -n 1")
			$final[$x][0] = $strComputer[$x][0]
			$final[$x][1] = $objItem.Started
			$final[$x][2] = $output
		Next
	Else
		$output = _GetDOSOutput("ping " & $strComputer[$x][0] & " -n 1")
		$final[$x][0] = $strComputer[$x][0]
		$final[$x][1] = "Server did not respond to WMI"
		$final[$x][2] = $output
	EndIf
	$query = 'insert into [Report$] ("Server","RDP","Ping") values (' & "'" & $final[$x][0] & "','" & $final[$x][1] & "','" & $final[$x][2] & "')"
	$conn.execute($query)
Next
;~ _ArrayDisplay($final)

$conn.close
$conn = ""

#region Format Excel
$oExcel = ObjCreate("Excel.Application")
$oExcel.WorkBooks.Open($report)
$oExcel.Application.DisplayAlerts = 0
$oExcel.Visible = 0
$iLastUsedCol = $oExcel.ActiveWorkbook.Sheets("Report").UsedRange.Columns.Count
$iLastUsedRow = $oExcel.ActiveWorkbook.Sheets("Report").UsedRange.Rows.Count
$range = $oExcel.Activesheet.Cells($iLastUsedRow, $iLastUsedCol).Address
$range = StringReplace($range, "$", "")
$range = "A1:" & $range
$oExcel.Range($range).Select
$oExcel.Selection.FormatConditions.Add($xlExpression, "", "=MOD(ROW(),2)=0", "")
$oExcel.Selection.FormatConditions($oExcel.Selection.FormatConditions.Count).SetFirstPriority
With $oExcel.Selection.FormatConditions(1).Interior
	.PatternColorIndex = $xlAutomatic
	.ThemeColor = $xlThemeColorDark1
	.TintAndShade = -0.14996795556505
EndWith
$oExcel.Selection.FormatConditions(1).StopIfTrue = True
With $oExcel.Selection.Borders
	.LineStyle = $xlContinuous
	.Weight = $xlThin
	.ColorIndex = $xlAutomatic
EndWith
$oExcel.Range("A1:C1").Select
With $oExcel.Selection.Interior
	.Pattern = $xlSolid
	.PatternColorIndex = $xlAutomatic
	.ThemeColor = $xlThemeColorDark1
	.TintAndShade = -0.499984740745262
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
	.ThemeColor = $xlThemeColorDark1
	.TintAndShade = 0
	.ThemeFont = $xlThemeFontMinor
EndWith
$oExcel.Columns("A:C").AutoFit
$oExcel.Range("A1").Select
$oExcel.ActiveWorkBook.Save
$oExcel.Quit()
$oExcel = 0
#endregion Format Excel

ConsoleWrite("End Time: " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF)
;~ FileClose($logFile)
Local $logFile = FileOpen(@ScriptDir & "\Error.log", 0)
$string = FileRead ($logFile,1)
FileClose($logFile)
If $string <> "" Then
    MsgBox (0,"","Some errors were generated. Please review the log file. If you have any servers that did not respond to RDP this is likely the issue.")
EndIf

MsgBox(0,"","Complete")

Func _GetDOSOutput($sCommand)
	Local $iPID, $sOutput = ""

	$iPID = Run('"' & @ComSpec & '" /c ' & $sCommand, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	While 1
		$sOutput &= StdoutRead($iPID, False, False)
		If @error Then
			ExitLoop
		EndIf
		Sleep(10)
	WEnd
	If StringInStr($sOutput, "Ping request could not find host ") <> 0 Then
		$sOutput = "False"
	ElseIf StringInStr($sOutput, "Reply from ") <> 0 Then
		$sOutput = "True"
	Else
		$sOutput = "Unknown"
	EndIf
	Return $sOutput
EndFunc   ;==>_GetDOSOutput

Func MyErrFunc()
	;MsgBox(0,"","Error")
	Dim $oMyRet[2]
    $HexNumber = Hex($oMyError.number, 8)
    $oMyRet[0] = $HexNumber
    $oMyRet[1] = StringStripWS($oMyError.description, 3)
;~     MsgBox(0,"Error","### Com Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
	_FileWriteLog($logFile, "### Com Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
	_FileWriteLog($logFile, "  ")
	_FileWriteLog($logFile, "  ")
    SetError(1); something to check for when this function returns
    Return
EndFunc   ;==>MyErrFunc