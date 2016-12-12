#include <Date.au3>
#include <Excel.au3>
#include <GUIConstantsEx.au3>
$year = @YEAR & "|" & @YEAR + 1 & "|" & @YEAR + 2 & "|" & @YEAR + 3 & "|" & @YEAR + 4 & "|" & @YEAR + 5
$file = @ScriptDir & "\tacschedule.xls"
$msg = 0

GUICreate("TAC schedule", 240, 100)
GUISetBkColor(0x094F95)
$button = GUICtrlCreateButton("Submit", 70, 70, 100, 30, 0x4000)
GUICtrlCreateLabel("Year needed:", 20, 10, 100, 20, 0x1000)
GUICtrlSetColor(-1, 0xF0F0E0)
GUICtrlCreateLabel("Team on shift Jan 1:", 20, 40, 100, 20, 0x1000)
GUICtrlSetColor(-1, 0xF0F0E0)
$combo1 = GUICtrlCreateCombo("", 140, 10, 80)
GUICtrlSetData(-1, $year, @YEAR + 1)
$combo2 = GUICtrlCreateCombo("", 140, 40, 80)
GUICtrlSetData(-1, "Team 1|Team 3")
GUISetState()

While $msg <> $GUI_EVENT_CLOSE
	$msg = GUIGetMsg()
	Select
		Case $msg = $button
			$year = GUICtrlRead($combo1)
			$team = GUICtrlRead($combo2)
			$first = $year & "/01/01"
			$count = _DateDiff('d', $first, $year & "/12/31")
			If $team = "Team 3" Then
				$team2 = "Team 4"
			Else
				$team2 = "Team 2"
			EndIf
			Main()
			MsgBox(0, "", "The file has been created " & @LF & @LF & $file)
			Exit
		Case $msg = $GUI_EVENT_CLOSE
			Exit
	EndSelect
WEnd

Func Main()
	$oExcel = ObjCreate("Excel.Application")
	$oExcel.Application.DisplayAlerts = 0
	$oExcel.Visible = 0
	$oExcel.WorkBooks.Add
	$oExcel.ActiveSheet.Name = "Schedule"
	$oExcel.Activesheet.Cells(1, 1).Value = "Title"
	$oExcel.Activesheet.Cells(1, 2).Value = "Start Time"
	$oExcel.Activesheet.Cells(1, 3).Value = "End Time"
	$oExcel.ActiveWorkBook.SaveAs($file, -4143)
	$oExcel.Quit()

	$conn = ObjCreate("ADODB.Connection")
	$DSN = ("Driver={Microsoft Excel Driver (*.xls)};DBQ=" & $file & ";readOnly=false;Nullable=true")
	$conn.Open($DSN)

	For $x = 0 To $count
		$day = _DateToDayOfWeekISO(StringMid($first, 1, 4), StringMid($first, 6, 2), StringMid($first, 9, 2))
		If $x <> 0 Then
			If $day = 1 or $day = 3 Or $day = 5 Then
				If $team = "Team 1" Then
					$team = "Team 3"
					$team2 = "Team 4"
				ElseIf $team = "Team 3" Then
					$team = "Team 1"
					$team2 = "Team 2"
				EndIf
			EndIf
		EndIf

		$start = StringMid($first, 6, 2) & "/" & StringMid($first, 9, 2) & "/" & StringMid($first, 1, 4) & " 07:00 AM"
		$end = StringMid($first, 6, 2) & "/" & StringMid($first, 9, 2) & "/" & StringMid($first, 1, 4) & " 07:00 PM"
		$query = 'insert into [Schedule$] ("Title","Start Time","End Time") values (' & "'" & $team & "','" & $start & "','" & $end & "')"
		$conn.execute($query)

		$start = StringMid($first, 6, 2) & "/" & StringMid($first, 9, 2) & "/" & StringMid($first, 1, 4) & " 07:00 PM"
		$first = _DateAdd('D', "1", $first)
		$end = StringMid($first, 6, 2) & "/" & StringMid($first, 9, 2) & "/" & StringMid($first, 1, 4) & " 07:00 AM"
		$query = 'insert into [Schedule$] ("Title","Start Time","End Time") values (' & "'" & $team2 & "','" & $start & "','" & $end & "')"
		$conn.execute($query)
	Next
	$conn.Close
EndFunc   ;==>Main
