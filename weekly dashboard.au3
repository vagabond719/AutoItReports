#region Includes
#include <Array.au3>
#include "EzMySql.au3"
#include <Date.au3>
#include <OutlookEx.au3>
#include <Constants.au3>
#endregion Includes
#region Variable Declaration
$today = _DateToDayOfWeekISO(@YEAR, @MON, @MDAY)
$today = "-" & $today
$weekend = _DateAdd('d', $today, @YEAR & "/" & @MON & "/" & @MDAY)
$today = $today - 7
$weekstart = _DateAdd('d', $today, @YEAR & "/" & @MON & "/" & @MDAY)
$weekend = StringReplace($weekend, "/", "-")
$weekstart = StringReplace($weekstart, "/", "-")
$daterange = "remediatedon >= '" & $weekstart & " 00:00:00' and remediatedon <= '" & $weekend & " 23:59:59'"
$daterange2 = "timestamp >= '" & $weekstart & " 00:00:00' and timestamp <= '" & $weekend & " 23:59:59'"
$to = "email"
$message = "Weekly Dashboard Report: " & $weekstart & " - " & $weekend & "<br><br>Please contact Josh Boley with any questions or concerns.<br><br>email"
#endregion Variable Declaration

#region Create MYSQL Connections
If Not _EzMySql_Startup() Then
	MsgBox(0, "Error Starting MySql", "Error: " & @error & @CR & "Error string: " & _EzMySql_ErrMsg())
	Exit
EndIf

If Not _EzMySql_Open("", "server", "password", "DB", "port") Then
	MsgBox(0, "Error opening Database", "Error: " & @error & @CR & "Error string: " & _EzMySql_ErrMsg())
	Exit
EndIf
#endregion Create MYSQL Connections
#region Create report.xls create ODB connection
$file = @ScriptDir & "\report.xls"
If FileExists($file) Then
	FileDelete($file)
EndIf
$oExcel = ObjCreate("Excel.Application")
$oExcel.Application.DisplayAlerts = 0
$oExcel.Visible = 0
$oExcel.WorkBooks.Add
$oExcel.ActiveSheet.Name = "Report"
$oExcel.ActiveWorkbook.Sheets("Sheet2").Select()
$oExcel.ActiveWorkbook.Sheets("Sheet2").Delete()
$oExcel.ActiveWorkbook.Sheets("Sheet3").Select()
$oExcel.ActiveWorkbook.Sheets("Sheet3").Delete()
$oExcel.ActiveWorkBook.SaveAs($file, -4143)
$oExcel.Quit()

$conn = ObjCreate("ADODB.Connection")
$DSN = ("Driver={Microsoft Excel Driver (*.xls)};DBQ=" & $file & ";readOnly=false;imex=1;HDR=Yes")
$conn.Open($DSN)

$query = 'Create table AnalystResponse(Analyst integer,"Total Time" integer, "Ticket Count" integer, AverageResponse integer)'
$conn.execute($query)

$query = 'Create table ClearedAlerts(AlertType text(50),AlertCount integer, TotalTicket integer)'
$conn.execute($query)

$query = 'Create table ExpandedAlertType(AlertType text(50),Description text(250), TotalTicket integer)'
$conn.execute($query)

$query = 'Create table Hourly(Hours text(250),Sunday integer, Monday integer, Tuesday integer, Wednesday integer, Thursday integer, Friday integer, Saturday integer)'
$conn.execute($query)

$query = 'Create table TicketsCreated(SSO integer, "Assigned and Remediated" integer, "Assigned but not Remediated" integer, "Remediated Only" integer, "Total alerts with tickets" integer,' & _
		'"Total Remediated" integer)'
$conn.execute($query)

$query = 'Create table TopOffenders(Device text(250),Alerts integer, AlertType text(250), Description text(250))'
$conn.execute($query)

$query = 'Create table PingTest(Server text(250), Response text(250))'
$conn.execute($query)

$query = 'Create table ResponseTimeRawData(AlertID integer, "Received Date" date, "Assigned Date" date, "Assigned SSO" integer, "Remediated Date" date, "Remediated By" integer, "Minutes till assigned" integer, "Minutes till remediated" integer, "Response Time" integer, Analyst integer)'
$conn.execute($query)

#endregion Create report.xls create ODB connection
#region Create Analyst array. ;This array is used to generate reports for each individual.
$analyst = _EzMySql_GetTable2d("select distinct(assignedto) from Alerts where " & $daterange & " and assignedto <> ''")
$error = @error
If Not IsArray($analyst) Then MsgBox(0, " error", $error)
#endregion Create Analyst array. ;This array is used to generate reports for each individual.
#region Main
TicketsCreated()
ResponseTime()
ResponseTimeProcessing()
ClearedAlerts()
ExpandedAlertType()
TopOffenders()
Hourly()
PingTest()
$conn.close
$conn = ""
$DSN = ""
$RS = ""

$oExcel = ObjCreate("Excel.Application")
$oExcel.Application.DisplayAlerts = 0
$oExcel.Visible = 0
$oExcel.WorkBooks.Open($file)
$oExcel.ActiveWorkbook.Sheets("Report").Select()
$oExcel.ActiveWorkbook.Sheets("Report").Delete()
Formating("AnalystResponse")
Formating("ClearedAlerts")
Formating("ExpandedAlertType")
Formating("TicketsCreated")
Formating("ResponseTimeRawData")
Formating("TopOffenders")
Formating("Hourly")
Formating("PingTest")
$oExcel.ActiveWorkBook.SaveAs($file, -4143)
$oExcel.Quit()
$outlook = _OL_Open()
_OL_Wrapper_SendMail($outlook, $to, "", "", "Weekly Dashboard Report: " & $weekstart & " - " & $weekend, $message, $file, $olFormatHTML, $olImportanceNormal)
_OL_Close($outlook)
#endregion Main

Func TicketsCreated()

	$rows = UBound($analyst)
	Dim $report[$rows + 1][6]
	$y = 1
	$report[0][0] = "SSO"
	$report[0][1] = "Assigned and Remediated"
	$report[0][2] = "Assigned but not Remediated"
	$report[0][3] = "Remediated Only"
	$report[0][4] = "Total alerts with tickets"
	$report[0][5] = "Total Remediated"
	$report[$rows][0] = "000000"

	For $x = 1 To UBound($analyst) - 1
		If $analyst[$x][0] = 'System' Then
			$query = "select Count(ticketnumber) from Alerts where " & $daterange & " and assignedto = '" & $analyst[$x][0] & "'"
			$ticketcount = _EzMySql_GetTable2d($query) ;create a list of analysts to use to query their individual performance
			$report[$y][0] = $analyst[$x][0]
			$report[$y][1] = $ticketcount[1][0]
			$report[$rows][1] = $report[$rows][1] + $report[$y][1]
			$y += 1
		ElseIf StringLen($analyst[$x][0]) <> 9 Then
			ContinueLoop
		Else
			$report[$y][0] = $analyst[$x][0]
			$query = "select Count(distinct(ticketnumber)) from Alerts where " & $daterange & " and ticketnumber not like '111%' and " & _
					"ticketnumber not like '000%' and ticketnumber <> '' and assignedto = '" & $analyst[$x][0] & "' and remediatedby = '" & $analyst[$x][0] & "'"
			$ticketcount = _EzMySql_GetTable2d($query) ;Count of unique tickets where assignedto and remediated by are the same analyst
			$report[$y][1] = $ticketcount[1][0]

			$report[$rows][1] = $report[$rows][1] + $report[$y][1]
			$query = "select Count(distinct(ticketnumber)) from Alerts where " & $daterange & " and ticketnumber not like '111%' and " & _
					"ticketnumber not like '000%' and ticketnumber <> '' and assignedto = '" & $analyst[$x][0] & "' and remediatedby <> '" & $analyst[$x][0] & "'"
			$ticketcount = _EzMySql_GetTable2d($query) ;Count of unique tickets where assignedto and remediated by are not the same analyst and remediated by the analyst
			$report[$y][2] = $ticketcount[1][0]
			$report[$rows][2] = $report[$rows][2] + $report[$y][2]

			$query = "select Count(distinct(ticketnumber)) from Alerts where " & $daterange & " and ticketnumber not like '111%' and " & _
					"ticketnumber not like '000%' and ticketnumber <> '' and remediatedby = '" & $analyst[$x][0] & "'"
			$ticketcount = _EzMySql_GetTable2d($query)
			$report[$y][3] = $ticketcount[1][0]
			$report[$rows][3] = $report[$rows][3] + $report[$y][3]

			$query = "select Count(ticketnumber) from Alerts where " & $daterange & " and ticketnumber not like '111%' and " & _
					"ticketnumber not like '000%' and ticketnumber <> '' and remediatedby = '" & $analyst[$x][0] & "'"
			$ticketcount = _EzMySql_GetTable2d($query)
			$report[$y][4] = $ticketcount[1][0]
			$report[$rows][4] = $report[$rows][4] + $report[$y][4]


			$query = "select Count(remediatedby) from Alerts where " & $daterange & " and remediatedby = '" & $analyst[$x][0] & "'"
			$ticketcount = _EzMySql_GetTable2d($query)
			$report[$y][5] = $ticketcount[1][0]
			$report[$rows][5] = $report[$rows][5] + $report[$y][5]
			$query = 'Insert into [TicketsCreated$] (SSO, "Assigned and Remediated", "Assigned but not Remediated", "Remediated Only", "Total alerts with tickets", "Total Remediated")' & _
					" values (" & $report[$y][0] & "," & $report[$y][1] & "," & $report[$y][2] & "," & $report[$y][3] & "," & $report[$y][4] & "," & _
					$report[$y][5] & ")"
			$conn.execute($query)
			$y += 1
		EndIf
	Next
	$query = 'Insert into [TicketsCreated$] (SSO, "Assigned and Remediated", "Assigned but not Remediated", "Remediated Only", "Total alerts with tickets", "Total Remediated")' & _
			" values (" & $report[$rows][0] & "," & $report[$rows][1] & "," & $report[$rows][2] & "," & $report[$rows][3] & "," & $report[$rows][4] & "," & _
			$report[$rows][5] & ")"
	$conn.execute($query)
EndFunc   ;==>TicketsCreated

Func ResponseTime()
;~ 	$array = _EzMySql_GetTable2d("select * from activitylog where action <> 'Update' and alertid in (select alertid from Alerts where " & $daterange & " and alerttype not in ('Network','Splunk')) order by alertid, timestamp")
	$array = _EzMySql_GetTable2d("select * from activitylog where alertid in (select alertid from Alerts where " & $daterange & " and alerttype not in ('Network','Splunk')) order by alertid, timestamp")
	$error = @error
	If Not IsArray($array) Then MsgBox(0, " error", $error)
;~ 	_ArrayDisplay($array)

	Dim $array2[1][10]
	$counter = -1
	For $x = 0 To UBound($array) - 1
		$array[$x][3] = StringReplace($array[$x][3], "-", "/")
		If $array[$x][5] = "Received" Then
			$counter += 1
			ReDim $array2[$counter + 1][10]
			$array2[$counter][0] = $array[$x][1]
			$array2[$counter][1] = $array[$x][3]
		ElseIf $array[$x][5] = "Assigned" And $array[$x][1] = $array[$x - 1][1] Then
			$array2[$counter][2] = $array[$x][3]
			$array2[$counter][3] = $array[$x][2]
		ElseIf $array[$x][5] = "Update" And $array[$x][1] = $array[$x - 1][1] And $array2[$counter][2] = "" Then
			$array2[$counter][2] = $array[$x][3]
			$array2[$counter][3] = $array[$x][2]
		ElseIf $array[$x][5] = "Remediated" Then
			$array2[$counter][4] = $array[$x][3]
			$array2[$counter][5] = $array[$x][2]
		Else
			$array[$x][6] = "GenFault"
		EndIf
	Next

	For $x = 0 To UBound($array2) - 1
		If $array2[$x][2] <> "" Then
			$array2[$x][6] = _DateDiff('n', $array2[$x][1], $array2[$x][2])
		EndIf
		$array2[$x][7] = _DateDiff('n', $array2[$x][1], $array2[$x][4])
		If $array2[$x][3] <> "" Then
			$array2[$x][8] = $array2[$x][6]
			$array2[$x][9] = $array2[$x][3]
		Else
			$array2[$x][8] = $array2[$x][7]
			$array2[$x][9] = $array2[$x][5]
		EndIf
		If $array2[$x][2] <> "" Then
			$query = 'Insert into [ResponseTimeRawData$] (AlertID, "Received Date", "Assigned Date", "Assigned SSO", "Remediated Date", "Remediated By", "Minutes till assigned", ' & _
					'"Minutes till remediated", "Response Time", Analyst) values (' & $array2[$x][0] & ",'" & $array2[$x][1] & "','" & $array2[$x][2] & "'," & $array2[$x][3] & ",'" & _
					$array2[$x][4] & "'," & $array2[$x][5] & "," & $array2[$x][6] & "," & $array2[$x][7] & "," & $array2[$x][8] & "," & $array2[$x][9] & ")"
			$conn.execute($query)
		Else
			$query = 'Insert into [ResponseTimeRawData$] (AlertID, "Received Date", "Remediated Date", "Remediated By", ' & _
					'"Minutes till remediated", "Response Time", Analyst) values (' & $array2[$x][0] & ",'" & $array2[$x][1] & "','" & _
					$array2[$x][4] & "'," & $array2[$x][5] & "," & $array2[$x][7] & "," & $array2[$x][8] & "," & $array2[$x][9] & ")"
;~ 			If StringInStr($query,"''") <> 0 Then
;~ 				MsgBox(0,"",$query)
;~ 			EndIf
			$conn.execute($query)
		EndIf
	Next
EndFunc   ;==>ResponseTime

Func ResponseTimeProcessing()
	$RS = ""
	$RS = ObjCreate("ADODB.Recordset")
	$query = 'select distinct("Analyst") from [ResponseTimeRawData$]'
	$RS.open($query, $conn)
	$array = $RS.GetRows()
	Dim $results[UBound($array) + 1][4]
	$RS = ""
	For $x = 0 To UBound($array) - 1
		$RS = ObjCreate("ADODB.Recordset")
		$query = 'select sum("Response Time"), count(Analyst), sum("Response Time")/count(Analyst) from [ResponseTimeRawData$] where Analyst=' & $array[$x][0]
		$RS.open($query, $conn)
		$array2 = $RS.GetRows()
		$RS = ""
		$results[$x][0] = $array[$x][0]
		$results[$x][1] = $array2[0][0]
		$results[$x][2] = $array2[0][1]
		$results[$x][3] = $array2[0][2]
	Next

	$RS = ObjCreate("ADODB.Recordset")
	$query = 'select sum("Response Time"), count(Analyst), sum("Response Time")/count(Analyst) from [ResponseTimeRawData$]'
	$RS.open($query, $conn)
	$array2 = $RS.GetRows()
	$results[$x][0] = "'000000'"
	$results[$x][1] = $array2[0][0]
	$results[$x][2] = $array2[0][1]
	$results[$x][3] = $array2[0][2]
	For $x = 0 To UBound($results) - 1
		$query = 'insert into [AnalystResponse$] (Analyst ,"Total Time" , "Ticket Count" , AverageResponse) values (' & $results[$x][0] & "," & $results[$x][1] & "," & _
				$results[$x][2] & "," & $results[$x][3] & ")"
		$conn.execute($query)
	Next
EndFunc   ;==>ResponseTimeProcessing

Func ClearedAlerts()
	$array = _EzMySql_GetTable2d("SELECT alerttype, count(*) as AlertCount  FROM alerts where assignedto <> 'System' and " & $daterange & " group by alerttype having count(*)>0")

	$array2 = _EzMySql_GetTable2d("SELECT alerttype, count(*) as TotalTickets FROM alerts where assignedto <> 'System' and " & $daterange & " and ticketnumber <> " & '""' & _
			" and ticketnumber not like '111%' and ticketnumber not like '000%' group by alerttype having count(*)>0")

	Dim $array3[UBound($array)][3]

	$z = 0

	For $x = 0 To UBound($array) - 1
		$array3[$z][0] = $array[$x][0]
		$array3[$z][1] = $array[$x][1]
		For $y = 0 To UBound($array2) - 1
			If $array[$x][0] = $array2[$y][0] Then
				$array3[$z][2] = $array2[$y][1]
			EndIf
			If $array3[$z][2] = "" Then
				$array3[$z][2] = 0
			EndIf
		Next
		$z += 1
	Next
	For $x = 1 To UBound($array3) - 1
		$query = 'insert into [ClearedAlerts$](AlertType ,AlertCount , TotalTicket) values (' & "'" & $array3[$x][0] & "'," & $array3[$x][1] & "," & $array3[$x][2] & ")"
		$conn.execute($query)
	Next
EndFunc   ;==>ClearedAlerts

Func ExpandedAlertType()
	Dim $report[1][3]
	$row = 0

	$array = _EzMySql_GetTable2d("SELECT distinct(alerttype) FROM alerts where " & $daterange)

	For $x = 1 To UBound($array) - 1
		$array2 = _EzMySql_GetTable2d("select distinct(description), count(description) from Alerts where " & $daterange & " and assignedto='' and alerttype='" & $array[$x][0] & "' and (ticketnumber like '111%' or ticketnumber like '000%' or ticketnumber = '') group by description")
		For $y = 1 To UBound($array2) - 1
			$report[$row][0] = $array[$x][0]
			$report[$row][1] = $array2[$y][0]
			$report[$row][2] = $array2[$y][1]
			$row += 1
			ReDim $report[$row + 1][3]
		Next
	Next
	For $x = 0 To UBound($report) - 1
		If $report[$x][0] <> "" Then
			$query = 'insert into [ExpandedAlertType$](AlertType ,Description , TotalTicket) values (' & "'" & $report[$x][0] & "','" & $report[$x][1] & "'," & $report[$x][2] & ")"
			$conn.execute($query)
		EndIf
	Next
EndFunc   ;==>ExpandedAlertType

Func Formating($sheet)
	$oExcel.ActiveWorkbook.Sheets($sheet).Select()
	$oExcel.Worksheets($sheet).Columns("A:I").AutoFit
	$oExcel.Range("A1").Entirerow.Font.Bold = True
	If $sheet = "AnalystResponse" Or $sheet = "MonthlyMetrics" Or $sheet = "TicketsCreated" Or $sheet = "Hourly" Then
		$iLastRow = $oExcel.ActiveSheet.UsedRange.rows.Count ; last row in any column
		$oExcel.Activesheet.Cells($iLastRow, 1).Value = "Totals"
		$oExcel.Activesheet.Cells($iLastRow, 1).Font.Bold = True
	EndIf
EndFunc   ;==>Formating

Func TopOffenders()
	$query = "select device, count(device) as alerts, alerttype, description from Alerts where " & $daterange & " group by device having count(device) > 20 order by alerts desc"
	$array = _EzMySql_GetTable2d($query)

	For $x = 1 To UBound($array) - 1
		$query = "insert into [TopOffenders$](Device ,Alerts , AlertType , Description) values ('" & $array[$x][0] & "'," & $array[$x][1] & ",'" & $array[$x][2] & "','" & $array[$x][3] & "')"
		$conn.execute($query)
	Next
EndFunc   ;==>TopOffenders

Func Hourly()
	$query = "select substring(timestamp,9,5) as Sub, count(alertid) from activitylog where " & $daterange2 & " and action='Received' group by sub having count(alertid) > 0"
	$array = _EzMySql_GetTable2d($query)

	Dim $array2[26][9]
	$array2[0][0] = "Time"
	$array2[0][1] = "Sunday"
	$array2[0][2] = "Monday"
	$array2[0][3] = "Tuesday"
	$array2[0][4] = "Wednesday"
	$array2[0][5] = "Thursday"
	$array2[0][6] = "Friday"
	$array2[0][7] = "Saturday"
	$array2[25][0] = "Totals"
	$array2[1][0] = "00:00"
	$array2[2][0] = "01:00"
	$array2[3][0] = "02:00"
	$array2[4][0] = "03:00"
	$array2[5][0] = "04:00"
	$array2[6][0] = "05:00"
	$array2[7][0] = "06:00"
	$array2[8][0] = "07:00"
	$array2[9][0] = "08:00"
	$array2[10][0] = "09:00"
	$array2[11][0] = "10:00"
	$array2[12][0] = "11:00"
	$array2[13][0] = "12:00"
	$array2[14][0] = "13:00"
	$array2[15][0] = "14:00"
	$array2[16][0] = "15:00"
	$array2[17][0] = "16:00"
	$array2[18][0] = "17:00"
	$array2[19][0] = "18:00"
	$array2[20][0] = "19:00"
	$array2[21][0] = "20:00"
	$array2[22][0] = "21:00"
	$array2[23][0] = "22:00"
	$array2[24][0] = "23:00"

	$row = 1
	$column = 1
	For $x = 1 To UBound($array) - 1
		$array2[$row][$column] = $array[$x][1]
		$array2[25][$column] = $array2[25][$column] + $array[$x][1]
		$row += 1
		If $row = 25 Then
			$row = 1
			$column += 1
		EndIf
	Next
	For $x = 1 To UBound($array2) - 1
		If $array2[$x][0] = "" Then
			$array2[$x][0] = 0
		EndIf
		If $array2[$x][1] = "" Then
			$array2[$x][1] = 0
		EndIf
		If $array2[$x][2] = "" Then
			$array2[$x][2] = 0
		EndIf
		If $array2[$x][3] = "" Then
			$array2[$x][3] = 0
		EndIf
		If $array2[$x][4] = "" Then
			$array2[$x][4] = 0
		EndIf
		If $array2[$x][5] = "" Then
			$array2[$x][5] = 0
		EndIf
		If $array2[$x][6] = "" Then
			$array2[$x][6] = 0
		EndIf
		If $array2[$x][7] = "" Then
			$array2[$x][7] = 0
		EndIf
		$query = "insert into [Hourly$](Hours ,Sunday , Monday , Tuesday , Wednesday , Thursday , Friday , Saturday) values ('" & $array2[$x][0] & "'," & $array2[$x][1] & "," & $array2[$x][2] & _
				"," & $array2[$x][3] & "," & $array2[$x][4] & "," & $array2[$x][5] & "," & $array2[$x][6] & "," & $array2[$x][7] & ")"
		$conn.execute($query)
	Next
EndFunc   ;==>Hourly

Func PingTest()
	$servers = _EzMySql_GetTable2d("select distinct(device) from Alerts where " & $daterange & " and description in ('COMMUNICATION ERROR','SYSTEM IS UNREACHABLE')")
	Dim $results[UBound($servers)][2]

	For $x = 1 To UBound($servers) - 1
		$results[$x][0] = $servers[$x][0]
		$output = _GetDOSOutput("ping " & $servers[$x][0] & " -n 1") & @CRLF
		If StringInStr($output, "Ping request could not find host ") <> 0 Then
			$results[$x][1] = "False"
		ElseIf StringInStr($output, "Reply from ") <> 0 Then
			$results[$x][1] = "True"
		Else
			$results[$x][1] = "Unknown"
		EndIf
		$query = "insert into [PingTest$] (server,response) values ('" & $results[$x][0] & "','" & $results[$x][1] & "')"
		$conn.execute($query)
	Next
;~ 	_ArrayDisplay($results)
EndFunc   ;==>PingTest

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
	Return $sOutput
EndFunc   ;==>_GetDOSOutput