#region includes
#include <Excel.au3>
#include <Array.au3>
#include <Date.au3>
#include <Parse.au3>
#include <IE.au3>
#include <File.au3>
#include <OutlookEx.au3>
#include <ServiceLoad.au3>
#endregion includes
_IEErrorHandlerRegister()
#region define holiday
Dim $holiday[21][2]
$holiday[1][0] = "New Year's Day"
$holiday[2][0] = "Martin Luther King"
$holiday[3][0] = "Presidents' Day"
$holiday[4][0] = "Memorial Day"
$holiday[5][0] = "Independence Day"
$holiday[6][0] = "Labor Day"
$holiday[7][0] = "Columbus Day"
$holiday[8][0] = "Veterans Day"
$holiday[9][0] = "Thanksgiving Day"
$holiday[10][0] = "Christmas"
$holiday[11][0] = "Last New Year's Day"
$holiday[12][0] = "Last Martin Luther King"
$holiday[13][0] = "Last Presidents' Day"
$holiday[14][0] = "Last Memorial Day"
$holiday[15][0] = "Last Independence Day"
$holiday[16][0] = "Last Labor Day"
$holiday[17][0] = "Last Columbus Day"
$holiday[18][0] = "Last Veterans Day"
$holiday[19][0] = "Last Thanksgiving Day"
$holiday[20][0] = "Last Christmas"

$mlk = _DateToDayOfWeekISO(@YEAR, "01", "15")
If $mlk <> 1 Then
	$mlk = @YEAR & "/01/" & 15 + (8 - $mlk)
ElseIf $mlk = 1 Then
	$mlk = @YEAR & "/01/15"
EndIf

$lmlk = _DateToDayOfWeekISO((@YEAR - 1), "01", "15")
If $lmlk <> 1 Then
	$lmlk = (@YEAR - 1) & "/01/" & 15 + (8 - $lmlk)
ElseIf $lmlk = 1 Then
	$lmlk = (@YEAR - 1) & "/01/15"
EndIf

$pres = _DateToDayOfWeekISO(@YEAR, "02", "15")
If $pres <> 1 Then
	$pres = @YEAR & "/02/" & 15 + (8 - $pres)
ElseIf $pres = 1 Then
	$pres = @YEAR & "02/15"
EndIf

$lpres = _DateToDayOfWeekISO((@YEAR - 1), "02", "15")
If $lpres <> 1 Then
	$lpres = (@YEAR - 1) & "/02/" & 15 + (8 - $lpres)
ElseIf $lpres = 1 Then
	$lpres = (@YEAR - 1) & "02/15"
EndIf

$memor = _DateToDayOfWeekISO(@YEAR, "05", "31")
If $memor <> 1 Then
	$memor = @YEAR & "/05/" & 32 - $memor
ElseIf $memor = 1 Then
	$memor = @YEAR & "/05/31"
EndIf

$lmemor = _DateToDayOfWeekISO((@YEAR - 1), "05", "31")
If $lmemor <> 1 Then
	$lmemor = (@YEAR - 1) & "/05/" & 32 - $lmemor
ElseIf $lmemor = 1 Then
	$lmemor = (@YEAR - 1) & "/05/31"
EndIf

$lab = _DateToDayOfWeekISO(@YEAR, "09", "01")
If $lab <> 1 Then
	$lab = @YEAR & "/09/" & 1 + (8 - $lab)
ElseIf $lab = 1 Then
	$lab = @YEAR & "/09/01"
EndIf

$llab = _DateToDayOfWeekISO((@YEAR - 1), "09", "01")
If $llab <> 1 Then
	$llab = (@YEAR - 1) & "/09/" & 1 + (8 - $llab)
ElseIf $llab = 1 Then
	$llab = (@YEAR - 1) & "/09/01"
EndIf

$columbus = _DateToDayOfWeekISO(@YEAR, "10", "08")
If $columbus <> 1 Then
	$columbus = @YEAR & "/10/" & 8 + (8 - $columbus)
ElseIf $columbus = 1 Then
	$columbus = @YEAR & "/10/08"
EndIf

$lcolumbus = _DateToDayOfWeekISO((@YEAR - 1), "10", "08")
If $lcolumbus <> 1 Then
	$lcolumbus = (@YEAR - 1) & "/10/" & 8 + (8 - $lcolumbus)
ElseIf $lcolumbus = 1 Then
	$lcolumbus = (@YEAR - 1) & "/10/08"
EndIf

$thank = _DateToDayOfWeekISO(@YEAR, "11", "22")
If $thank <> 4 Then
	$thank = @YEAR & "/11/" & 22 + (8 - $thank)
ElseIf $thank = 4 Then
	$thank = @YEAR & "/11/22"
EndIf

$lthank = _DateToDayOfWeekISO((@YEAR - 1), "11", "22")
If $lthank <> 4 Then
	$lthank = (@YEAR - 1) & "/11/" & 22 + (8 - $lthank)
ElseIf $lthank = 4 Then
	$lthank = (@YEAR - 1) & "/11/22"
EndIf

$holiday[1][1] = @YEAR & "/01/01"
$holiday[2][1] = $mlk
$holiday[3][1] = $pres
$holiday[4][1] = $memor
$holiday[5][1] = @YEAR & "/07/04"
$holiday[6][1] = $lab
$holiday[7][1] = $columbus
$holiday[8][1] = @YEAR & "/11/11"
$holiday[9][1] = $thank
$holiday[10][1] = @YEAR & "/12/25"
$holiday[11][1] = (@YEAR - 1) & "/01/01"
$holiday[12][1] = $lmlk
$holiday[13][1] = $lpres
$holiday[14][1] = $lmemor
$holiday[15][1] = (@YEAR - 1) & "/07/04"
$holiday[16][1] = $llab
$holiday[17][1] = $lcolumbus
$holiday[18][1] = (@YEAR - 1) & "/11/11"
$holiday[19][1] = $lthank
$holiday[20][1] = (@YEAR - 1) & "/12/25"

$holiday[0][0] = UBound($holiday) - 1
#endregion define holiday
;~ #region Pull Email from Outlook
;~ $outlook = _OL_Open()
;~ $olookfolder = _OL_FolderAccess($outlook, "Functional.Idm@ge.com\Inbox")

;~ $emails = _OL_ItemFind($outlook, "Functional.Idm@ge.com\Inbox", $olMail, "", "Subject", "**Daily SLA Report**", "EntryID, Subject, SenderName, ReceivedTime", "[RecievedTime], True")

;~ _OL_ItemAttachmentSave($outlook, $emails[1][0], $olookfolder[3], 1, @ScriptDir & "\tsgnew.xls")

;~ _OL_ItemMove($outlook, $emails[1][0], $olookfolder[3], "Functional.Idm@ge.com\Inbox\Completed")
;~ #endregion Pull Email from Outlook
#region Create Blank TSGDATA.XLS file
$file = @ScriptDir & "\tsgdata.xls"
$oExcel = ObjCreate("Excel.Application")
$oExcel.Application.DisplayAlerts = 0
$oExcel.Visible = 0
$oExcel.WorkBooks.Add
$oExcel.ActiveSheet.Name = "Blank"
$oExcel.ActiveWorkbook.Sheets("Sheet2").Select()
$oExcel.ActiveWorkbook.Sheets("Sheet2").Delete()
$oExcel.ActiveWorkbook.Sheets("Sheet3").Select()
$oExcel.ActiveWorkbook.Sheets("Sheet3").Delete()
$oExcel.ActiveWorkBook.SaveAs($file, -4143)
$oExcel.Quit()
#endregion Create Blank TSGDATA.XLS file
#region Open ADO to TSGDATA.XLS file
$conn = ObjCreate("ADODB.Connection")
$RS = ObjCreate("ADODB.Recordset")
$DSN = ("Driver={Microsoft Excel Driver (*.xls)};DBQ=" & $file & ";readOnly=false;imex=1;HDR=Yes")
$conn.Open($DSN)

$file2 = @ScriptDir & "\tsgnew.xls"
$file3 = @ScriptDir & "\tsgold.xls"
#endregion Open ADO to TSGDATA.XLS file
#region create and populate tsgdata.xls file and !!calls SLA function!!
$query = 'select * into SRNew from [' & $file2 & '].[ServiceRequests$]'
$conn.execute($query)
$query = 'select * into INCNew from [' & $file2 & '].[Incidents$]'
$conn.execute($query)
$query = 'select * into SROld from [' & $file3 & '].[ServiceRequests$]'
$conn.execute($query)
$query = 'select * into INCOld from [' & $file3 & '].[Incidents$]'
$conn.execute($query)
$query = 'select * into ServiceRequests from [SRNew$] where "ADJ MADE/MISSED"=' & "'Made'"
$conn.execute($query)
$query = 'select * into Incidents from [INCNew$] where "ADJ MADE/MISSED"=' & "'Made'"
$conn.execute($query)
$query = 'select a.* into SRmiss from [SRNew$] a, [SROld$] b where a.id=b.id and a.closed=b.closed and a."ADJ MADE/MISSED"=' & "'Missed'"
$conn.execute($query)
$query = 'select a.* into INCmiss from [INCNew$] a, [INCOld$] b where a.id=b.id and a.closed=b.closed and a."ADJ MADE/MISSED"=' & "'Missed'"
$conn.execute($query)
$query = 'INSERT INTO [ServiceRequests$] (ID,Op5ened,Closed,"Exception time",TasktoTime,"Value",Name,Priority,"Business duration",OPENEDFW,CLOSEDFW,' & _
		'"Short description","Business Segment",Location,Sub,"Assignment group","Bus Days minus Exception","SLA Goal","Adj SLA Made","Adj SLA Missed",Day,' & _
		'"ADJ MADE/MISSED","RAW MADE","RAW MISSED","RAW MADE/MISSED",Reopened,"Close Code") select * from [SRmiss$]'
$conn.execute($query)
$query = 'INSERT INTO [Incidents$] (ID,Opened,Closed,"Exception time",TasktoTime,"Value",Name,Priority,"Business duration",OPENEDFW,CLOSEDFW,' & _
		'"Short description","Business Segment",Location,Sub,"Assignment group","Bus Dur minus Exception","SLA Goal","Adj SLA Made","Adj SLA Missed",Day,' & _
		'"ADJ MADE/MISSED","RAW MADE","RAW MISSED","RAW MADE/MISSED",Reopened,"Close Code") select * from [INCmiss$]'
$conn.execute($query)

$query = 'select a.id into Missed from [' & $file2 & '].[ServiceRequests$] a left outer join [' & $file3 & '].[ServiceRequests$] b on a.id = b.id where (a.closed<>b.closed or b.id is null) and a."ADJ MADE/MISSED"=' & "'Missed'"
$conn.execute($query)
$query = 'insert into [Missed$] (id) select a.id from [' & $file2 & '].[Incidents$] a left outer join [' & $file3 & '].[Incidents$] b on a.id = b.id where (a.closed<>b.closed or b.id is null) and a."ADJ MADE/MISSED"=' & "'Missed'"
$conn.execute($query)

Call("sla")

$query = 'select a."ID",a."Opened",a."Closed",b.exception as "Exception time",a."TasktoTime",a."Value",a."Name",a."Priority",a."Business duration",a."OPENEDFW",a."CLOSEDFW",' & _
		'a."Short description",a."Business Segment",a."Location",a."Sub",a."Assignment group",a."Bus Days minus Exception",a."SLA Goal",a."Adj SLA Made",a."Adj SLA Missed",' & _
		'a."Day",a."ADJ MADE/MISSED",a."RAW MADE",a."RAW MISSED",a."RAW MADE/MISSED",a."Reopened",a."Close Code" from  [SRNew$] a inner join [updates$] b on a.id=b.tickets'

$RS = ""
$RS = ObjCreate("ADODB.Recordset")
$RS.open($query, $conn)
$array = $RS.GetRows()
$row = UBound($array) - 1
For $x = 0 To $row ;Write updated SR values to ServiceRequests sheet
	If $array[$x][17] = 0 Then
		$sla = 7200
	ElseIf $array[$x][17] = 1 Then
		$sla = 14400
	ElseIf $array[$x][17] = 2 Then
		$sla = 86400
	ElseIf $array[$x][17] = 3 Then
		$sla = 259200
	ElseIf $array[$x][17] = 4 Then
		$sla = 432000
	EndIf
	$busdur = ($array[$x][8] - $array[$x][3])
	$array[$x][16] = ($array[$x][8] - $array[$x][3]) / 86400
	If $sla > $busdur Then
		$array[$x][18] = 1
		$array[$x][19] = 0
		$array[$x][21] = "Made"
	EndIf
	$query = 'INSERT INTO [ServiceRequests$] (ID,Opened,Closed,"Exception time",TasktoTime,"Value",Name,Priority,"Business duration",OPENEDFW,CLOSEDFW,' & _
			'"Short description","Business Segment",Location,Sub,"Assignment group","Bus Days minus Exception","SLA Goal","Adj SLA Made","Adj SLA Missed",Day,' & _
			'"ADJ MADE/MISSED","RAW MADE","RAW MISSED","RAW MADE/MISSED",Reopened,"Close Code") VALUES ' & _
			"('" & $array[$x][0] & "','" & $array[$x][1] & "','" & $array[$x][2] & "','" & $array[$x][3] & "','" & $array[$x][4] & _
			"','" & $array[$x][5] & "','" & $array[$x][6] & "','" & $array[$x][7] & "','" & $array[$x][8] & "','" & $array[$x][9] & "','" & _
			$array[$x][10] & "','" & $array[$x][11] & "','" & $array[$x][12] & "','" & $array[$x][13] & "','" & $array[$x][14] & "','" & _
			$array[$x][15] & "','" & $array[$x][16] & "','" & $array[$x][17] & "','" & $array[$x][18] & "','" & $array[$x][19] & "','" & $array[$x][20] & _
			"','" & $array[$x][21] & "','" & $array[$x][22] & "','" & $array[$x][23] & "','" & $array[$x][24] & "','" & $array[$x][25] & "','" & $array[$x][26] & "');"
	$conn.execute($query)
Next
;------------------------------
$RS = ""
$RS = ObjCreate("ADODB.Recordset")

$query = 'select a."ID",a."Opened",a."Closed",b.exception as "Exception time",a."TasktoTime",a."Value",a."Name",a."Priority",a."Business duration",a."OPENEDFW",a."CLOSEDFW",' & _
		'a."Short description",a."Business Segment",a."Location",a."Sub",a."Assignment group",a."Bus Dur minus Exception",a."SLA Goal",a."Adj SLA Made",a."Adj SLA Missed",a."Day",' & _
		'a."ADJ MADE/MISSED",a."RAW MADE",a."RAW MISSED",a."RAW MADE/MISSED",a."Reopened",a."Close Code" from  [INCNew$] a inner join [updates$] b on a.id=b.tickets'
$RS.open($query, $conn)
$array = $RS.GetRows()
$row = UBound($array) - 1
For $x = 0 To $row ;Write updated INC values to Incidents sheet
	If $array[$x][17] = 0 Then
		$sla = 7200
	ElseIf $array[$x][17] = 1 Then
		$sla = 14400
	ElseIf $array[$x][17] = 2 Then
		$sla = 86400
	ElseIf $array[$x][17] = 3 Then
		$sla = 259200
	ElseIf $array[$x][17] = 4 Then
		$sla = 432000
	EndIf
	$busdur = ($array[$x][8] - $array[$x][3])
	$array[$x][16] = ($array[$x][8] - $array[$x][3]) / 86400
	If $sla > $busdur Then
		$array[$x][18] = 1
		$array[$x][19] = 0
		$array[$x][21] = "Made"
	EndIf
	$query = 'INSERT INTO [Incidents$] (ID,Opened,Closed,"Exception time",TasktoTime,"Value",Name,Priority,"Business duration",OPENEDFW,CLOSEDFW,' & _
			'"Short description","Business Segment",Location,Sub,"Assignment group","Bus Dur minus Exception","SLA Goal","Adj SLA Made","Adj SLA Missed",Day,' & _
			'"ADJ MADE/MISSED","RAW MADE","RAW MISSED","RAW MADE/MISSED",Reopened,"Close Code") VALUES ' & _
			"('" & $array[$x][0] & "','" & $array[$x][1] & "','" & $array[$x][2] & "','" & $array[$x][3] & "','" & $array[$x][4] & _
			"','" & $array[$x][5] & "','" & $array[$x][6] & "','" & $array[$x][7] & "','" & $array[$x][8] & "','" & $array[$x][9] & "','" & _
			$array[$x][10] & "','" & $array[$x][11] & "','" & $array[$x][12] & "','" & $array[$x][13] & "','" & $array[$x][14] & "','" & _
			$array[$x][15] & "','" & $array[$x][16] & "','" & $array[$x][17] & "','" & $array[$x][18] & "','" & $array[$x][19] & "','" & $array[$x][20] & _
			"','" & $array[$x][21] & "','" & $array[$x][22] & "','" & $array[$x][23] & "','" & $array[$x][24] & "','" & $array[$x][25] & "','" & $array[$x][26] & "');"
	$conn.execute($query)
Next
#endregion create and populate tsgdata.xls file and !!calls SLA function!!
#region create and populate tsg.xls file
$file4 = @ScriptDir & "\tsg.xls"
$oExcel = ObjCreate("Excel.Application")
$oExcel.Application.DisplayAlerts = 0
$oExcel.Visible = 0
$oExcel.WorkBooks.Add

For $x = 1 To 2 ;Create
	If $x = 2 Then
		$table = "ServiceRequests"
		$columnname = "Bus Days minus Exception"
		$query = 'select * from [ServiceRequests$]'
	ElseIf $x = 1 Then
		$table = "Incidents"
		$columnname = "Bus Dur minus Exception"
		$query = 'select * from [Incidents$]'
	EndIf
	$RS = ""
	$RS = ObjCreate("ADODB.Recordset")
	$RS.open($query, $conn)
	$oExcel.ActiveWorkBook.WorkSheets.Add().Activate()
	$oExcel.ActiveSheet.Name = $table
	$oExcel.ActiveWorkBook.ActiveSheet.Cells(2, 1).CopyFromRecordset($RS)
	$oExcel.Activesheet.Cells(1, 1).Value = "ID"
	$oExcel.Activesheet.Cells(1, 2).Value = "Opened"
	$oExcel.Activesheet.Cells(1, 3).Value = "Closed"
	$oExcel.Activesheet.Cells(1, 4).Value = "Exception time"
	$oExcel.Activesheet.Cells(1, 5).Value = "TasktoTime"
	$oExcel.Activesheet.Cells(1, 6).Value = "Value"
	$oExcel.Activesheet.Cells(1, 7).Value = "Name"
	$oExcel.Activesheet.Cells(1, 8).Value = "Priority"
	$oExcel.Activesheet.Cells(1, 9).Value = "Business duration"
	$oExcel.Activesheet.Cells(1, 10).Value = "OPENEDFW"
	$oExcel.Activesheet.Cells(1, 11).Value = "CLOSEDFW"
	$oExcel.Activesheet.Cells(1, 12).Value = "Short description"
	$oExcel.Activesheet.Cells(1, 13).Value = "Business Segment"
	$oExcel.Activesheet.Cells(1, 14).Value = "Location"
	$oExcel.Activesheet.Cells(1, 15).Value = "Sub"
	$oExcel.Activesheet.Cells(1, 16).Value = "Assignment group"
	$oExcel.Activesheet.Cells(1, 17).Value = $columnname
	$oExcel.Activesheet.Cells(1, 18).Value = "SLA Goal"
	$oExcel.Activesheet.Cells(1, 19).Value = "Adj SLA Made"
	$oExcel.Activesheet.Cells(1, 20).Value = "Adj SLA Missed"
	$oExcel.Activesheet.Cells(1, 21).Value = "Day"
	$oExcel.Activesheet.Cells(1, 22).Value = "ADJ MADE/MISSED"
	$oExcel.Activesheet.Cells(1, 23).Value = "RAW MADE"
	$oExcel.Activesheet.Cells(1, 24).Value = "RAW MISSED"
	$oExcel.Activesheet.Cells(1, 25).Value = "RAW MADE/MISSED"
	$oExcel.Activesheet.Cells(1, 26).Value = "Reopened"
	$oExcel.Activesheet.Cells(1, 27).Value = "Close Code"
	$oExcel.Columns("d:d").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("e:e").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("i:i").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("j:j").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("k:k").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("q:q").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("r:r").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("s:s").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("t:t").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("w:w").Select
	$oExcel.Selection.TextToColumns
	$oExcel.Columns("x:x").Select
	$oExcel.Selection.TextToColumns
Next

$oExcel.ActiveWorkbook.Sheets("Sheet1").Select()
$oExcel.ActiveWorkbook.Sheets("Sheet1").Delete()
$oExcel.ActiveWorkbook.Sheets("Sheet2").Select()
$oExcel.ActiveWorkbook.Sheets("Sheet2").Delete()
$oExcel.ActiveWorkbook.Sheets("Sheet3").Select()
$oExcel.ActiveWorkbook.Sheets("Sheet3").Delete()
$oExcel.ActiveWorkBook.SaveAs($file4, -4143)
$oExcel.Quit()
#endregion create and populate tsg.xls file
Call("mailer")
#region close connections
$conn.close
$conn = ""
$RS = ""
$DSN = ""
FileDelete($file)
FileDelete($file2)
FileDelete($file3)
FileMove($file4, $file3)
#endregion close connections
Func SLA() ;Generate list of missed tickets to loop through and calculate exception !!Calls QuitIE, Service, excep Functions!!
	$query = "CREATE TABLE Updates (Tickets varchar(255),Exception integer,Update varchar(255))"
	$conn.execute($query)
	#region create log file
	If FileExists(@ScriptDir & "\log.csv") Then
		$filedate = FileGetTime(@ScriptDir & "\log.csv",1)
		$filedate = $filedate[0] & "/" & $filedate[1] & "/" & $filedate[2]
		If _NowCalcDate() <> $filedate Then
			FileDelete(@ScriptDir & "\log.csv")
		EndIf
	EndIf
	$log = FileOpen(@ScriptDir & "/log.csv", 1)
	FileClose($log)
	$textconn = ObjCreate("ADODB.Connection")
	$textRS = ObjCreate("ADODB.Recordset")
	$textDSN = ("Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & @ScriptDir & "\")
	$textconn.Open($textDSN)
	$textquery = "select * from log.csv"
	$textRS.open($textquery, $textconn)
	$comparray = $textRS.GetRows()
	$textconn.close
	$textRS = ""
	$textDSN = ""
	$textconn = ""
	#endregion create log file
	;_arraydisplay($comparray)
	#region write log file values to tsgdata.xls
	For $i = 1 To UBound($comparray)-1
		If $comparray[$i][0] <> "" Then
			$query = 'insert into [Updates$] ("Tickets","Exception","Update") values ' & "('" & $comparray[$i][0] & "'," & $comparray[$i][1] & ",'Processed')"
			$conn.execute($query)
		EndIf
	Next
	#endregion write log file values to tsgdata.xls

	Do
		$RS = ObjCreate("ADODB.Recordset")
		$query = "select a.id from [Missed$] a left outer join [Updates$] b on a.id=b.Tickets where b.Update is null"
		$RS.open($query, $conn, 3, 3, 0x0001)

		$array = $RS.GetRows()

		If Not IsArray($array) Then
			ExitLoop
		EndIf

		$arraycount = UBound($array)
		$arraycount2 = UBound($array, 2)

		If $arraycount > 100 Then
			ReDim $array[100][$arraycount2]
		EndIf

		Call("QuitIE")

		$site = Call("Service")

		For $x = 0 To (UBound($array) - 1)
			#region handle Service Now Failures and bad statuses
			$totalex = 0
			$exception = 0

			$frame = _IEFrameGetCollection($site, 0)
			If @error <> 0 Then
				Do
					Call("QuitIE")
					$site = Call("Service")
					$frame = _IEFrameGetCollection($site, 0)
				Until IsObj($frame)
			EndIf

			$form = _IEGetObjByName($site, "sysparm_search")
			$search = _IEGetObjByName($site, "textsearch")
			_IEFormElementSetValue($form, $array[$x][0])
			_IEFormSubmit($search)

			_ServiceLoad($site, $frame)

			$site = _IEAttach($array[$x][0], "title", 1)

			$frame = _IEFrameGetCollection($site, 0)

			$check = StringInStr($array[$x][0], "SR")

			$ticketcheck = 1
			If $check > 0 Then
				$ticketcheck = 2
				$ticket = "u_service_request"
			Else
				$ticket = "incident"
			EndIf

			$statustest = _IEGetObjByName($frame, $ticket & ".state")
			$statusvalue = _IEFormElementGetValue($statustest)

			If $ticket = "incident" Then
				If $statusvalue <> "2" And $statusvalue <> "3" Then
					$frame1 = _IEFrameGetObjByName($site, "gsft_nav")
					If @error <> 0 Then
						Do
							Call("QuitIE")
							$site = Call("Service")
							$frame1 = _IEFrameGetCollection($site, "gsft_nav")
						Until IsObj($frame1)
					EndIf

					$form = _IEGetObjById($frame1, "4aeebcd20a0a0b9a00572ae3ad68b072")
					_IEAction($form, "click")

					$query = 'insert into [Updates$] ("Tickets","Exception","Update") values ' & "('" & $array[$x][0] & "',0,'Incorrect Ticket Status')"
					$conn.execute($query)
					$log = FileOpen(@ScriptDir & "/log.csv", 1)
					FileWrite($log, $array[$x][0] & @CRLF)
					FileClose($log)
					ContinueLoop
				EndIf
			ElseIf $ticket = "u_service_request" Then
				If $statusvalue <> "4" And $statusvalue <> "3" Then
					$frame1 = _IEFrameGetObjByName($site, "gsft_nav")
					If @error <> 0 Then
						Do
							Call("QuitIE")
							$site = Call("Service")
							$frame1 = _IEFrameGetCollection($site, "gsft_nav")
						Until IsObj($frame1)
					EndIf

					$form = _IEGetObjById($frame1, "4aeebcd20a0a0b9a00572ae3ad68b072")
					_IEAction($form, "click")

					$query = 'insert into [Updates$] ("Tickets","Exception","Update") values ' & "('" & $array[$x][0] & "',0,'Incorrect Ticket Status')"
					$conn.execute($query)
					$log = FileOpen(@ScriptDir & "/log.csv", 1)
					FileWrite($log, $array[$x][0] & @CRLF)
					FileClose($log)
					ContinueLoop
				EndIf
			EndIf
			#endregion handle Service Now Failures and bad statuses

			$totalex = Call("Excep", $frame)

			#region handle Successful Exceptions
			$totaldays = Int($totalex / 86400)
			$totalhours = Int(($totalex - ($totaldays * 86400)) / 3600)
			$totalmins = Int(($totalex - ($totalhours * 3600) - ($totaldays * 86400)) / 60)
			$totalsecs = Int($totalex - ($totalmins * 60) - ($totalhours * 3600) - ($totaldays * 86400))


			$exception = $totaldays & " " & $totalhours & ":" & $totalmins & ":" & $totalsecs

			$statusobj = _IEGetObjByName($frame, $ticket & ".state")
			$status = _IEFormElementGetValue($statusobj)

			If $status = 3 Then
				$frame1 = _IEFrameGetObjByName($site, "gsft_nav")
				If @error <> 0 Then
					Do
						Call("QuitIE")
						$site = Call("Service")
						$frame1 = _IEFrameGetCollection($site, "gsft_nav")
					Until IsObj($frame1)
				EndIf
				$form = _IEGetObjById($frame1, "4aeebcd20a0a0b9a00572ae3ad68b072")
				_IEAction($form, "click")
				$log = FileOpen(@ScriptDir & "/log.csv", 1)
				FileWrite($log, $array[$x][0] & "|" & $totalex & @CRLF)
				FileClose($log)

				$query = 'insert into [Updates$] ("Tickets","Exception","Update") values ' & "('" & $array[$x][0] & "'," & $totalex & ",'Manual Update')"
				$conn.execute($query)
				ContinueLoop
			Else
				$class = _IEGetObjByName($frame, $ticket & ".u_classification")
				$classcheck = _IEFormElementGetValue($class)

				If $classcheck = "" Then
					_IEFormElementOptionSelect($class, "Monitoring")
				EndIf

				$l1solv = _IEGetObjByName($frame, $ticket & ".u_solvable_l1")
				$l1solvcheck = _IEFormElementGetValue($l1solv)

				If $l1solvcheck = "" Then
					_IEFormElementOptionSelect($l1solv, "No")
				EndIf

				$endu = _IEGetObjByName($frame, $ticket & ".u_verified_solution_with_end_u")
				$enducheck = _IEFormElementGetValue($endu)

				If $enducheck = "" Then
					_IEFormElementOptionSelect($endu, "No")
				EndIf

				If $check = 0 Then
					$closecode = _IEGetObjByName($frame, $ticket & ".close_code")
					$closecodecheck = _IEFormElementGetValue($closecode)

					If $closecodecheck = "" Then
						_IEFormElementOptionSelect($closecode, "Resolved Remotely")
					EndIf
				Else
					$closecode = _IEGetObjByName($frame, $ticket & ".u_close_code")
					$closecodecheck = _IEFormElementGetValue($closecode)

					If $closecodecheck = "" Then
						_IEFormElementOptionSelect($closecode, "Request Completed")
					EndIf
				EndIf

				$subclass = _IEGetObjByName($frame, $ticket & ".u_sub_classification")
				$subclasscheck = _IEFormElementGetValue($subclass)

				If $subclasscheck = "" Then
					Do
						$subclass = _IEGetObjByName($frame, $ticket & ".u_sub_classification")
						$subclasscheck = _IEFormElementOptionSelect($subclass, "Monitoring")
						Sleep(100)
					Until $subclasscheck = 1
				EndIf

				$closenotes = _IEGetObjByName($frame, $ticket & ".close_notes")
				$closenotescheck = _IEFormElementGetValue($closenotes)

				If $closenotescheck = "" Then
					_IEFormElementOptionSelect($subclass, "Updated by Audit")
				EndIf

				If $ticket = "incident" Then
					$vendor = _IEGetObjByName($frame, $ticket & ".u_vendor_involved")
					$vendorcheck = _IEFormElementOptionSelect($vendor, "No")
				EndIf

				$closedby = _IEGetObjByName($frame, $ticket & ".closed_by")
				$closedbycheck = _IEFormElementGetValue($closedby)

				If $closedbycheck = "" Then
					_IEFormElementSetValue($closedby, "b3d99e580a0a3cbd006263b236e7efde")
				EndIf

				$addcom = _IEGetObjByName($frame, $ticket & ".u_additional_closing_workgroup")
				_IEFormElementSetValue($addcom, "SLA Audit Complete")


				$excode = _IEGetObjByName($frame, $ticket & ".u_exception_code")
				_IEFormElementOptionSelect($excode, "Awaiting approval for shared data access")
				$exceptionval = _IEGetObjByName($frame, $ticket & ".u_exception_time")
				_IEFormElementSetValue($exceptionval, $exception)

				$save = _IEGetObjByName($frame, "sysverb_update")
				_IEAction($save, "click")

				_ServiceLoad($site, $frame)

				$log = FileOpen(@ScriptDir & "/log.csv", 1)
				FileWrite($log, $array[$x][0] & "|" & $totalex & @CRLF)
				FileClose($log)

				$query = 'insert into [Updates$] ("Tickets","Exception","Update") values ' & "('" & $array[$x][0] & "'," & $totalex & ",'Updated')"
				$conn.execute($query)

			EndIf
			#endregion handle Successful Exceptions
		Next
		$RS = ""
	Until Not IsArray($array)
EndFunc   ;==>SLA
Func Service() ;Handles ServiceNow login
	ShellExecute("iexplore.exe", "about:blank")
	WinWait("Blank Page")
	$oIE = _IEAttach("about:blank", "url")
	_IELoadWait($oIE)
	_IENavigate($oIE, "https://getsg.service-now.com/")
	_IELoadWait($oIE)

	Do
		$oIE = _IEAttach("", "instance")
		$check1 = _IEPropertyGet($oIE, "locationurl")
		$check = StringInStr($check1, "ssologin.corporate.ge.com")
		$check2 = StringInStr($check1, "getsg.service-now")
		Sleep(100)
	Until $check > 0 Or $check2 > 0

	If $check > 0 Then

		Do
			$name = _IEGetObjByName($oIE, "username")
			$nameset = _IEFormElementSetValue($name, "502117737")

			$pass = _IEGetObjByName($oIE, "password")
			$passset = _IEFormElementSetValue($pass, "SSOLogin12@T")
		Until $nameset = 1 And $passset = 1


		$submit = _IEGetObjByName($oIE, "Submit")
		_IEAction($submit, "click")

		_IELoadWait($oIE)

		_IENavigate($oIE, "https://getsg.service-now.com/")

		_IELoadWait($oIE)

		_IEAction($oIE, "refresh")

		Return ($oIE)

	Else
		Return ($oIE)
	EndIf
EndFunc   ;==>Service
Func QuitIE() ;Kills IE

	Local $iearray[1]
	$iearray[0] = 0

	Local $i = 1
	While 1
		$iein = _IEAttach("", "instance", $i)
		If @error = $_IEStatus_NoMatch Then ExitLoop
		ReDim $iearray[$i + 1]
		$iearray[$i] = $iein
		$iearray[0] = $i
		$i += 1
	WEnd

	For $x = 1 To $iearray[0]
		_IEQuit($iearray[$x])
	Next
EndFunc   ;==>QuitIE
Func Excep($frame) ;Parse Notes Table Calculates Exception And Returns Exception Time !!Calls Time and Holiday Functions!!
	$table = _IETableGetCollection($frame)
	Local $iNumTables = @extended

	For $x = 0 To $iNumTables
		;MsgBox(0,"", $x & " to " & $iNumTables)

		$table = _IETableGetCollection($frame, $x)

		$exarray = _IETableWriteToArray($table, True)
		If StringMid($exarray[0][0], 3, 1) = "-" Then
			If StringMid($exarray[0][0], 6, 1) = "-" Then
				;_ArrayDisplay($exarray)
				ExitLoop
			EndIf
		EndIf
	Next

	For $x = 1 To UBound($exarray) - 1
		$checkgrp = StringInStr($exarray[$x][0], "Assignment Group: ")
		$checkst = StringInStr($exarray[$x][0], "Status: ")
		If $checkgrp <> 0 And $checkst = 0 Then
			$pos1 = StringInStr($exarray[$x][0], "Category: ")
			$pos3 = StringInStr($exarray[$x][0], "was: ")

			If $pos3 >= $pos1 Then
				$pos1 = $pos3
			EndIf

			$pos2 = $pos1 - $checkgrp

			$exarray[$x - 1][2] = StringMid($exarray[$x][0], $checkgrp + 18, $pos2 - 18)
			$exarray[$x - 1][1] = 1
			$date = StringLeft($exarray[$x - 1][0], 19)
			$year = StringMid($date, 7, 4)
			$mon = StringLeft($date, 2)
			$day = StringMid($date, 4, 2)
			$hms = StringRight($date, 8)

			$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			$exarray[$x - 1][5] = "GROUP"

		ElseIf $checkst <> 0 And $checkgrp = 0 Then
			If StringInStr($exarray[$x][0], "Status: Completed") <> 0 Then
				$exarray[$x - 1][2] = "Completed"
				$exarray[$x - 1][1] = 1
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Resolved") <> 0 Then
				$exarray[$x - 1][2] = "Resolved"
				$exarray[$x - 1][1] = 1
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Pending") <> 0 Then
				$exarray[$x - 1][2] = "Pending"
				$exarray[$x - 1][1] = 1
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Pending Authorization") <> 0 Then
				$exarray[$x - 1][2] = "Pending Authorization"
				$exarray[$x - 1][1] = 1
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Pending User Info") <> 0 Then
				$exarray[$x - 1][2] = "Pending User Info"
				$exarray[$x - 1][1] = 1
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Reopened") <> 0 Then
				$exarray[$x - 1][2] = "Reopened"
				$exarray[$x - 1][1] = 1
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Active") <> 0 Then
				$exarray[$x - 1][2] = "Active"
				$exarray[$x - 1][1] = 1
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: In Progress") <> 0 Then
				$exarray[$x - 1][2] = "In Progress"
				$exarray[$x - 1][1] = 1
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: End-user Updated") <> 0 Then
				$exarray[$x - 1][2] = "End-User Updated"
				$exarray[$x - 1][1] = 1
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")
			EndIf

			$exarray[$x - 1][5] = "STATUS"

		ElseIf $checkst <> 0 And $checkgrp <> 0 Then
			$pos1 = StringInStr($exarray[$x][0], "Category: ")
			$pos3 = StringInStr($exarray[$x][0], "was: ")

			If $pos3 >= $pos1 Then
				$pos1 = $pos3
			EndIf

			$pos2 = $pos1 - $checkgrp

			$exarray[$x - 1][2] = StringMid($exarray[$x][0], $checkgrp + 18, $pos2 - 18)
			$date = StringLeft($exarray[$x - 1][0], 19)
			$year = StringMid($date, 7, 4)
			$mon = StringLeft($date, 2)
			$day = StringMid($date, 4, 2)
			$hms = StringRight($date, 8)

			$exarray[$x - 1][3] = $year & $mon & $day & StringReplace($hms, ":", "")

			$exarray[$x - 1][1] = 1

			If StringInStr($exarray[$x][0], "Status: Completed") <> 0 Then
				$exarray[$x - 1][3] = "Completed"
				$exarray[$x - 1][1] = 2
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][4] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Resolved") <> 0 Then
				$exarray[$x - 1][3] = "Resolved"
				$exarray[$x - 1][1] = 2
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][4] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Pending") <> 0 Then
				$exarray[$x - 1][3] = "Pending"
				$exarray[$x - 1][1] = 2
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][4] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Pending Authorization") <> 0 Then
				$exarray[$x - 1][3] = "Pending Authorization"
				$exarray[$x - 1][1] = 2
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][4] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Pending User Info") <> 0 Then
				$exarray[$x - 1][3] = "Pending User Info"
				$exarray[$x - 1][1] = 2
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][4] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Reopened") <> 0 Then
				$exarray[$x - 1][3] = "Reopened"
				$exarray[$x - 1][1] = 2
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][4] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: Active") <> 0 Then
				$exarray[$x - 1][3] = "Active"
				$exarray[$x - 1][1] = 2
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][4] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: In Progress") <> 0 Then
				$exarray[$x - 1][3] = "In Progress"
				$exarray[$x - 1][1] = 2
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][4] = $year & $mon & $day & StringReplace($hms, ":", "")
			ElseIf StringInStr($exarray[$x][0], "Status: End-user Updated") <> 0 Then
				$exarray[$x - 1][3] = "End -User Updated"
				$exarray[$x - 1][1] = 2
				$date = StringLeft($exarray[$x - 1][0], 19)
				$year = StringMid($date, 7, 4)
				$mon = StringLeft($date, 2)
				$day = StringMid($date, 4, 2)
				$hms = StringRight($date, 8)

				$exarray[$x - 1][4] = $year & $mon & $day & StringReplace($hms, ":", "")
			EndIf

		EndIf
	Next

	;_arraydisplay($exarray)

	Local $exarray2[2][6]
	$count = 1
	For $z = 1 To UBound($exarray) - 1
		If $exarray[$z][1] = 1 Then
			ReDim $exarray2[$count + 1][6]
			If $exarray[$z][5] = "STATUS" Then
				$exarray2[$count][0] = "STATUS"
			Else
				$exarray2[$count][0] = "GROUP"
			EndIf
			$exarray2[$count][1] = $exarray[$z][2]
			$exarray2[$count][2] = $exarray[$z][3]
			$count += 1
		ElseIf $exarray[$z][1] = 2 Then
			ReDim $exarray2[$count + 2][6]
			$exarray2[$count][0] = "GROUP"
			$exarray2[$count][1] = $exarray[$z][2]
			$exarray2[$count][2] = $exarray[$z][4]
			$exarray2[$count + 1][0] = "STATUS"
			$exarray2[$count + 1][1] = $exarray[$z][3]
			$exarray2[$count + 1][2] = $exarray[$z][4]
			$count += 1
		EndIf
		$exarray2[0][0] = UBound($exarray2) - 1
	Next

	$redimrow = UBound($exarray2)
	$redimcol = UBound($exarray2, 2)

	Dim $exarray3[$redimrow][$redimcol]

	$redimcount = UBound($exarray2) - 1

	For $z = 0 To UBound($exarray2) - 1
		If $z <> 0 Then
			$exarray3[$z - 1][3] = $exarray2[$redimcount][2]
		EndIf
		If $exarray2[$redimcount][0] = "GROUP" Then
			If StringInStr($exarray2[$redimcount][1], @CR) <> 0 Then
				$exarray2[$redimcount][1] = StringLeft($exarray2[$redimcount][1], StringInStr($exarray2[$redimcount][1], @CR) - 1)
			EndIf
			If StringInStr($exarray2[$redimcount][1], " was: ") <> 0 Then
				$exarray2[$redimcount][1] = StringLeft($exarray2[$redimcount][1], StringInStr($exarray2[$redimcount][1], " was: "))
			EndIf
			$exarray3[$z][0] = $exarray3[$z - 1][0]
			$exarray3[$z][1] = $exarray3[$z - 1][1]
			$exarray3[$z][2] = $exarray2[$redimcount][2]
			$exarray3[$z][5] = $exarray2[$redimcount][1]
			If $exarray3[$z - 1][5] = "" Then
				$exarray3[$z - 1][5] = $exarray3[$z][5]
			EndIf
			$redimcount -= 1
			ContinueLoop
		EndIf

		If $exarray2[$redimcount][0] = "STATUS" Then
			$exarray3[$z][0] = $exarray2[$redimcount][0]
			$exarray3[$z][1] = $exarray2[$redimcount][1]
			$exarray3[$z][2] = $exarray2[$redimcount][2]
			$exarray3[$z][5] = $exarray2[$redimcount][5]
			If $exarray3[$z][5] = "" And $z <> 0 Then
				$exarray3[$z][5] = $exarray3[$z - 1][5]
			EndIf
			$redimcount -= 1
		EndIf
		If $exarray3[$z][0] = "" Then
			$exarray3[$z][0] = "Total Exception"
		EndIf
	Next

	$arrayend = UBound($exarray3) - 1
	For $z = 0 To UBound($exarray3) - 1
		If $exarray3[$z][2] = $exarray3[$z][3] And $z <> $arrayend Then
			ContinueLoop
		EndIf
		If $exarray3[$z][3] = "" Then
			ContinueLoop
		EndIf
		If StringInStr($exarray3[$z][5], "L2 TSG Account Admin Support") = 0 Then
			$exarray3[$z][4] = Call("Time", $exarray3[$z][2], $exarray3[$z][3])
			$exarray3[$z][4] = $exarray3[$z][4] - Call("Holiday", $exarray3[$z][2], $exarray3[$z][3])
		EndIf
		If StringInStr($exarray3[$z][5], "L2 TSG Account Admin Support") <> 0 And StringInStr($exarray3[$z][1], "Pending") <> 0 Then
			$exarray3[$z][4] = Call("Time", $exarray3[$z][2], $exarray3[$z][3])
			$exarray3[$z][4] = $exarray3[$z][4] - Call("Holiday", $exarray3[$z][2], $exarray3[$z][3])
		EndIf
		If (StringInStr($exarray3[$z][1], "Completed") <> 0 Or StringInStr($exarray3[$z][1], "Resolved") <> 0) And $exarray3[$z][3] <> "" Then
			$exarray3[$z][4] = Call("Time", $exarray3[$z][2], $exarray3[$z][3])
			$exarray3[$z][4] = $exarray3[$z][4] - Call("Holiday", $exarray3[$z][2], $exarray3[$z][3])
		EndIf
		If (StringInStr($exarray3[$z][1], "Resolved") <> 0 Or StringInStr($exarray3[$z][1], "Resolved") <> 0) And $exarray3[$z][3] <> "" Then
			$exarray3[$z][4] = Call("Time", $exarray3[$z][2], $exarray3[$z][3])
			$exarray3[$z][4] = $exarray3[$z][4] - Call("Holiday", $exarray3[$z][2], $exarray3[$z][3])
		EndIf
		If $exarray3[$z][4] < 0 Then
			$exarray3[$z][4] = 0
		EndIf
		$exarray3[$arrayend][4] = $exarray3[$arrayend][4] + $exarray3[$z][4]
	Next
	;________________________________
	;________________________________
	;________________________________
	;_ArrayDisplay($exarray2)
	;_ArrayDisplay($exarray3)
	Return ($exarray3[$arrayend][4])
EndFunc   ;==>Excep
Func Time($startdate, $enddate) ;Removes weekends from exception time
	$start = $startdate ;Open date of ticket
	$end = $enddate ;Close date of ticket
	$daycounter = "" ;Tracks the number of weekend days
	$adjustedstart = "" ;Time passed on start date
	$adjustedend = "" ;Time passed on End date
	;$datecatch = ""
	$differencestart = ((((StringMid($start, 9, 2)) * 3600) + (StringMid($start, 11, 2)) * 60) + StringMid($start, 13, 2)) ;Convert start time difference to sec
	$differneceend = 86400 - ((((StringMid($end, 9, 2)) * 3600) + (StringMid($end, 11, 2)) * 60) + StringMid($end, 13, 2)) ;Convert end time difference to sec
	$dayofstart = _DateToDayOfWeekISO(StringMid($start, 1, 4), StringMid($start, 5, 2), StringMid($start, 7, 2)) ;Check what day of the week the ticket was opened
	$dayofstart2 = _DateToDayOfWeekISO(StringMid($start, 1, 4), StringMid($start, 5, 2), StringMid($start, 7, 2)) ;Counter used when calculating the weekend days
	$dayofend = _DateToDayOfWeekISO(StringMid($end, 1, 4), StringMid($end, 5, 2), StringMid($end, 7, 2)) ;Check what day of the week the ticket was closeed
	$start = StringMid($start, 1, 4) & "/" & StringMid($start, 5, 2) & "/" & StringMid($start, 7, 2) & " " & _
			StringMid($start, 9, 2) & ":" & StringMid($start, 11, 2) & "" & ":" & StringMid($start, 13, 2) ;Convert format of date/time
	$end = StringMid($end, 1, 4) & "/" & StringMid($end, 5, 2) & "/" & StringMid($end, 7, 2) & " " & StringMid($end, 9, 2) & ":" & StringMid($end, 11, 2) & _
			"" & ":" & StringMid($end, 13, 2) ;Convert format of date/time
	$startcalc = StringMid($start, 1, 10) & " 00:00:00" ;Set start time to the begining of the day
	$endcalc = StringMid($end, 1, 10) & " 23:59:59" ;Set end time to the end of the day
	$diffday = _DateDiff('d', $startcalc, $endcalc) ;Determine the number of days between opened and closed
	$diffsec = _DateDiff('s', $start, $end)

	For $y = 0 To $diffday
		If $dayofstart2 = 7 Then
			$daycounter += 1
			$dayofstart2 = 1
		ElseIf $dayofstart2 = 6 Then
			$daycounter += 1
			$dayofstart2 += 1
		Else
			$dayofstart2 += 1
		EndIf
	Next
	If $dayofstart > 5 Then
		$adjustedstart = $differencestart
	EndIf
	If $dayofend > 5 Then
		$adjustedend = $differneceend
	EndIf

	$daycounter = ($daycounter * 86400)
	$adjustedtime = $diffsec - $daycounter + $adjustedstart + $adjustedend
	;$array2[$y][3] = $adjustedtime
	Return ($adjustedtime)
EndFunc   ;==>Time
Func mailer() ;Emails final report
	$RS = ""
	$RS = ObjCreate("ADODB.Recordset")
	$query = 'Select sum("Adj SLA Made")/count("Adj SLA Made"), count("Adj SLA Made")-sum("Adj SLA Made"), sum("Adj SLA Made") FROM [ServiceRequests$]'
	$RS.open($query, $conn)
	$array = $RS.GetRows()
	$pct = $array[0][0]
	$pct = StringMid($array[0][0], 3, 3)
	$dec = StringMid($pct, 3, 1)
	$pct = StringMid($pct, 1, 2)
	$misssr = $array[0][1]
	$madesr = $array[0][2]
	If $dec > 5 Then
		$sr = $pct + 1
	Else
		$sr = $pct
	EndIf

	$RS = ""
	$RS = ObjCreate("ADODB.Recordset")
	$query = 'Select sum("Adj SLA Made")/count("Adj SLA Made"), count("Adj SLA Made")-sum("Adj SLA Made"), sum("Adj SLA Made") FROM [Incidents$]'
	$RS.open($query, $conn)
	$array = $RS.GetRows()
	$pct = ($array[0][0])
	$pct = StringMid($array[0][0], 3, 3)
	$dec = StringMid($pct, 3, 1)
	$pct = StringMid($pct, 1, 2)
	$missinc = $array[0][1]
	$madeinc = $array[0][2]
	If $dec > 5 Then
		$inc = $pct + 1
	Else
		$inc = $pct
	EndIf

	$conn.close

	$message = '<table border="1" width="60%">' & @CRLF & _
			'<col width="120">' & @CRLF & _
			'<col width="60">' & @CRLF & _
			'<col width="60">' & @CRLF & _
			'<col width="60">' & @CRLF & _
			'<tr bgcolor="#6699FF">' & @CRLF & _
			'<th> </center></th>' & @CRLF & _
			"<th><center>Missed</center></th>" & @CRLF & _
			"<th><center>Made</center></th>" & @CRLF & _
			"<th><center>SLA</center></th>" & @CRLF & _
			"</tr>" & @CRLF & _
			"<tr>" & @CRLF & _
			'<td bgcolor="#FFFFFF"><b>Service Requests</b></td>' & @CRLF & _
			"<td><center>" & $misssr & "</center></td>" & @CRLF & _
			"<td><center>" & $madesr & "</center></td>" & @CRLF & _
			"<td><center>" & $sr & "%</center></td>" & @CRLF & _
			"</tr>" & @CRLF & _
			"<tr>" & @CRLF & _
			'<td bgcolor="#FFFFFF"><b>Incidents</b></td>' & @CRLF & _
			"<td><center>" & $missinc & "</center></td>" & @CRLF & _
			"<td><center>" & $madeinc & "</center></td>" & @CRLF & _
			"<td><center>" & $inc & "%</center></td>" & @CRLF & _
			"</tr>" & @CRLF & _
			"</table>"

	$outlook = _OL_Open()
	_OL_Wrapper_SendMail($outlook, "502051460;Erwin.Hernandez@compucom.com; Paula.Bohne@compucom.com; Joe.Villanueva@compucom.com", "501967897", "502183923", "TSG - Daily Report - " & @MON & "/" & @MDAY & "/" & @YEAR, $message, @ScriptDir & "\tsg.xls", $olFormatHTML, $olImportanceNormal)
	;_OL_Wrapper_SendMail($outlook, "502051460", "", "", "TSG - Daily Report - " & @MON & "/" & @MDAY & "/" & @YEAR, $message, @ScriptDir & "\tsg.xls", $olFormatHTML, $olImportanceNormal)
	;502051460
	_OL_Close($outlook)
EndFunc   ;==>mailer
Func Holiday($hstartdate, $henddate) ;Removes holidays from exception time
	$hstart = $hstartdate ;Open date of ticket
	$hend = $henddate ;Close date of ticket
	$hdaycounter = 0 ;Tracks the number of weekend days
	$hadjustedstart = "" ;Time passed on start date
	$hadjustedend = "" ;Time passed on End date
	;$datecatch = ""
	$hdifferencestart = 86400 - ((((StringMid($hstart, 9, 2)) * 3600) + (StringMid($hstart, 11, 2)) * 60) + StringMid($hstart, 13, 2)) ;Convert start time difference to sec
	$hdifferneceend = ((((StringMid($hend, 9, 2)) * 3600) + (StringMid($hend, 11, 2)) * 60) + StringMid($hend, 13, 2)) ;Convert end time difference to sec
	$hstart = StringMid($hstart, 1, 4) & "/" & StringMid($hstart, 5, 2) & "/" & StringMid($hstart, 7, 2) & " " & _
			StringMid($hstart, 9, 2) & ":" & StringMid($hstart, 11, 2) & "" & ":" & StringMid($hstart, 13, 2) ;Convert format of date/time
	$hend = StringMid($hend, 1, 4) & "/" & StringMid($hend, 5, 2) & "/" & StringMid($hend, 7, 2) & " " & StringMid($hend, 9, 2) & ":" & StringMid($hend, 11, 2) & _
			"" & ":" & StringMid($hend, 13, 2) ;Convert format of date/time
	$hstartcalc = StringMid($hstart, 1, 10) & " 00:00:00" ;Set start time to the begining of the day
	$hendcalc = StringMid($hend, 1, 10) & " 23:59:59" ;Set end time to the end of the day
	$hdiffday = _DateDiff('d', $hstartcalc, $hendcalc) ;Determine the number of days between opened and closed
	$hdiffsec = _DateDiff('s', $hstart, $hend)

	$startval = $hstart
	$endval = StringLeft($hend, 10)
	For $q = 1 To $hdiffday + 1
		For $w = 1 To $holiday[0][0]
			If StringLeft($startval, 10) = $holiday[$w][1] And $startval = $hstart Then
				$hadjustedstart = $hdifferencestart
			ElseIf StringLeft($startval, 10) = $holiday[$w][1] And $startval <> $hstart And StringLeft($startval, 10) <> $endval Then
				$hdaycounter += 1
			ElseIf StringLeft($startval, 10) = $holiday[$w][1] And StringLeft($startval, 10) = $endval Then
				$hadjustedend = $hdifferneceend
			EndIf
		Next
		$startval = _DateAdd('D', "1", $startval)
	Next


	$hdaycounter = ($hdaycounter * 86400)
	$hadjustedtime = $hdaycounter + $hadjustedstart + $hadjustedend
	;$array2[$y][3] = $adjustedtime
	Return ($hadjustedtime)
EndFunc   ;==>Holiday